"""Generate monthly attendance Word documents and PDFs from the supplied templates."""

import os
import re
import tempfile
from copy import deepcopy
from contextlib import contextmanager
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from attendance_config import MONTHS, TEMPLATES, is_valid_buddhist_year
from docxtpl import DocxTemplate
from docx2pdf import convert
from lxml import etree


DATE_PATTERN = re.compile(
    r"(?P<month>มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|"
    r"สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)"
    r"\s*(?:พ\.ศ\.\s*)?(?P<year>[๐-๙0-9]{4})"
)


def _add_guard_date_dotted_underlines(paragraph, namespace):
    """Put dotted underlines under the guard sheet's month and year fields."""
    placeholder_pattern = re.compile(r"({{month}}|{{year}})")
    for run in list(paragraph.iter(f"{namespace}r")):
        text_nodes = list(run.iter(f"{namespace}t"))
        if len(text_nodes) != 1 or not text_nodes[0].text:
            continue
        text_node = text_nodes[0]
        parts = placeholder_pattern.split(text_node.text)
        if len(parts) == 1:
            continue

        parent = run.getparent()
        insert_at = parent.index(run)
        for part in parts:
            if not part:
                continue
            new_run = deepcopy(run)
            for child in list(new_run):
                if child.tag != f"{namespace}rPr":
                    new_run.remove(child)
            new_text = etree.SubElement(new_run, f"{namespace}t")
            new_text.text = part
            if part in {"{{month}}", "{{year}}"}:
                run_properties = new_run.find(f"{namespace}rPr")
                if run_properties is None:
                    run_properties = etree.SubElement(new_run, f"{namespace}rPr")
                underline = run_properties.find(f"{namespace}u")
                if underline is None:
                    underline = etree.SubElement(run_properties, f"{namespace}u")
                underline.set(f"{namespace}val", "dotted")
            if part[:1].isspace() or part[-1:].isspace():
                new_text.set(
                    "{http://www.w3.org/XML/1998/namespace}space", "preserve"
                )
            parent.insert(insert_at, new_run)
            insert_at += 1
        parent.remove(run)


@contextmanager
def initialized_com():
    """Initialize COM on the current thread for Microsoft Word automation."""
    import pythoncom

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


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
                        label_start = paragraph_text.find("วันที่")
                        leader_start = label_start + len("วันที่")
                        leader_text = paragraph_text[leader_start : match.start("month")]
                        if (
                            label_start >= 0
                            and leader_text
                            and not leader_text.strip()
                        ):
                            operations.append(
                                (
                                    leader_start,
                                    match.start("month"),
                                    "............ เดือน ",
                                )
                            )
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

                    if (
                        paragraph_text.lstrip().startswith("วันที่")
                        and "เดือน" not in paragraph_text
                    ):
                        _add_guard_date_dotted_underlines(paragraph, namespace)

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
    """Save a legacy .doc file as .docx through Microsoft Word."""
    from win32com.client import DispatchEx

    with initialized_com():
        word = DispatchEx("Word.Application")
        document = None
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(
                str(template_path.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            document.SaveAs2(str(docx_path.resolve()), FileFormat=16)
        finally:
            try:
                if document is not None:
                    document.Close(False)
            finally:
                word.Quit()


def convert_to_pdf(docx_path, pdf_path):
    """Convert a Word document to PDF with COM initialized on this thread."""
    with initialized_com():
        convert(str(docx_path), str(pdf_path))


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
        if is_valid_buddhist_year(year):
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
        convert(str(docx_path), str(pdf_path))
        print(f"Created PDF: {pdf_path}")


if __name__ == "__main__":
    main()
