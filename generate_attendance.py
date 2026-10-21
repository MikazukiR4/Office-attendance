"""Generate monthly attendance Word documents and PDFs from the supplied templates."""

import os
import re
import subprocess
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from attendance_config import MONTHS, TEMPLATES
from docxtpl import DocxTemplate
from lxml import etree


DATE_PATTERN = re.compile(
    r"(?P<month>มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|"
    r"สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)"
    r"\s*(?:พ\.ศ\.\s*)?(?P<year>[๐-๙0-9]{4})"
)


def prepare_template(template_path, prepared_path):
    """Replace only date text in the DOCX package, preserving its layout XML."""
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    xml_parts = []

    with ZipFile(template_path, "r") as source:
        for item in source.infolist():
            if item.filename == "word/document.xml" or re.fullmatch(
                r"word/(?:header|footer)\d+\.xml", item.filename
            ):
                root = etree.fromstring(source.read(item.filename))
                for paragraph in root.iter(f"{namespace}p"):
                    text_nodes = list(paragraph.iter(f"{namespace}t"))
                    paragraph_text = "".join(node.text or "" for node in text_nodes)

                    operations = []
                    for match in DATE_PATTERN.finditer(paragraph_text):
                        operations.extend(
                            (
                                (match.start("month"), match.end("month"), "{{month}}"),
                                (match.start("year"), match.end("year"), "{{year}}"),
                            )
                        )

                    for start, end, placeholder in sorted(operations, reverse=True):
                        node_starts = []
                        position = 0
                        for node in text_nodes:
                            node_starts.append(position)
                            position += len(node.text or "")

                        first = next(
                            (
                                index
                                for index, node_start in enumerate(node_starts)
                                if node_start <= start
                                < node_start + len(text_nodes[index].text or "")
                            ),
                            None,
                        )
                        last = next(
                            (
                                index
                                for index in range(len(text_nodes) - 1, -1, -1)
                                if node_starts[index] < end
                                <= node_starts[index] + len(text_nodes[index].text or "")
                            ),
                            None,
                        )
                        if first is None or last is None:
                            raise ValueError(
                                f"Could not locate date text in {item.filename}: "
                                f"{paragraph_text}"
                            )

                        first_text = text_nodes[first].text or ""
                        last_text = text_nodes[last].text or ""
                        first_offset = start - node_starts[first]
                        last_offset = end - node_starts[last]
                        if first == last:
                            text_nodes[first].text = (
                                first_text[:first_offset]
                                + placeholder
                                + first_text[last_offset:]
                            )
                        else:
                            text_nodes[first].text = (
                                first_text[:first_offset] + placeholder
                            )
                            for index in range(first + 1, last):
                                text_nodes[index].text = ""
                            text_nodes[last].text = last_text[last_offset:]

                xml_parts.append(
                    (
                        item.filename,
                        etree.tostring(
                            root,
                            encoding="UTF-8",
                            xml_declaration=True,
                            standalone=True,
                        ),
                    )
                )

        document_xml = b"\n".join(data for _, data in xml_parts)
        if b"{{month}}" not in document_xml or b"{{year}}" not in document_xml:
            raise ValueError(
                f"No month/year date or separate {{month}} and {{year}} placeholders "
                f"found in {template_path.name}"
            )

        with ZipFile(prepared_path, "w", compression=ZIP_DEFLATED) as destination:
            replaced = dict(xml_parts)
            for item in source.infolist():
                data = replaced[item.filename] if item.filename in replaced else source.read(item.filename)
                destination.writestr(item, data)


def convert_legacy_doc(template_path, docx_path):
    """Convert a legacy .doc file to .docx through WordCraft CLI."""
    run_wordcraft_convert(template_path, docx_path)


def convert_to_pdf(docx_path, pdf_path):
    """Convert a document to PDF through the headless WordCraft CLI."""
    run_wordcraft_convert(docx_path, pdf_path)


def run_wordcraft_convert(input_path, output_path):
    """Run WordCraft's headless converter and validate its output."""
    cli = os.environ.get("WORDCRAFT_CLI", "wordcraft-cli")
    try:
        result = subprocess.run(
            [
                cli,
                "convert",
                str(Path(input_path).resolve()),
                str(Path(output_path).resolve()),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as error:
        raise RuntimeError(
            "WordCraft CLI was not found. Install WordCraft and set "
            "WORDCRAFT_CLI to the wordcraft-cli executable."
        ) from error
    except subprocess.CalledProcessError as error:
        details = (error.stderr or error.stdout or "").strip()
        raise RuntimeError(
            f"WordCraft failed to convert {input_path} to {output_path}"
            + (f": {details}" if details else ".")
        ) from error

    if not Path(output_path).is_file():
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(
            f"WordCraft completed without creating {output_path}"
            + (f": {details}" if details else ".")
        )


def get_render_context():
    """Prompt for the month and Buddhist Era year."""
    while True:
        try:
            month = int(input("Enter target month (1-12): ").strip())
            if 1 <= month <= 12:
                break
        except ValueError:
            pass
        print("Please enter a month from 1 to 12.")

    while True:
        year = input("Enter Buddhist Era year (e.g., 2569): ").strip()
        if year.isascii() and year.isdigit() and len(year) == 4:
            break
        print("Please enter a four-digit Buddhist Era year, such as 2569.")

    return {"month": MONTHS[month - 1], "year": year}


def main():
    project_dir = Path(__file__).resolve().parent
    output_dir = project_dir / "output"
    os.makedirs(output_dir, exist_ok=True)
    context = get_render_context()
    generated_docs = []

    with tempfile.TemporaryDirectory() as temporary_dir:
        temporary_dir = Path(temporary_dir)
        for category, templates in TEMPLATES.items():
            for document_type, template_name in templates.items():
                template_path = project_dir / template_name
                if not template_path.is_file():
                    raise FileNotFoundError(f"Template not found: {template_path}")

                docx_template_path = template_path
                if template_path.suffix.lower() == ".doc":
                    docx_template_path = temporary_dir / f"{template_path.stem}.docx"
                    convert_legacy_doc(template_path, docx_template_path)

                prepared_path = temporary_dir / f"{category}_{document_type}.docx"
                prepare_template(docx_template_path, prepared_path)

                output_path = output_dir / (
                    f"{category}_{document_type}_{context['month']}_{context['year']}.docx"
                )
                template = DocxTemplate(str(prepared_path))
                template.render(context)
                template.save(str(output_path))
                generated_docs.append(output_path)
                print(f"Created Word document: {output_path}")

    for docx_path in generated_docs:
        pdf_path = docx_path.with_suffix(".pdf")
        convert_to_pdf(docx_path, pdf_path)
        print(f"Created PDF: {pdf_path}")


if __name__ == "__main__":
    main()
