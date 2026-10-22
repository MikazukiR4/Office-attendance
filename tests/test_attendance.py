import os
import shutil
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import pymupdf
from docxtpl import DocxTemplate
from lxml import etree

from app import (
    PREVIEW_CACHE_MAX_AGE_SECONDS,
    PREVIEW_CACHE_MAX_FILES,
    combine_pdf_bytes,
    get_template_signature,
    prune_preview_pdfs,
    save_preview_pdfs,
)
from attendance_config import TEMPLATES, is_valid_buddhist_year
from generate_attendance import (
    convert_legacy_doc,
    convert_to_pdf,
    enforce_attendance_font_size,
    prepare_template,
)


NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WORDCRAFT_CLI = os.environ.get(
    "WORDCRAFT_CLI",
    str(Path(__file__).resolve().parents[1] / ".tools" / "wordcraft" / "wordcraft-cli.exe"),
)


class AttendanceFontSizeTests(unittest.TestCase):
    """The attendance sheets must show every visible run at 16 pt on one page."""

    @classmethod
    def setUpClass(cls):
        cls.project_dir = Path(__file__).resolve().parents[1]

    def render_attendance_sheet(self, category, directory):
        """Render one attendance sheet and return the final DOCX path."""
        template_path = self.project_dir / TEMPLATES[category]["Table"]
        docx_template_path = template_path
        if template_path.suffix.lower() == ".doc":
            docx_template_path = Path(directory) / "converted.docx"
            convert_legacy_doc(template_path, docx_template_path)

        prepared_path = Path(directory) / "prepared.docx"
        prepare_template(docx_template_path, prepared_path)

        raw_path = Path(directory) / "raw.docx"
        final_path = Path(directory) / "final.docx"
        template = DocxTemplate(str(prepared_path))
        template.render({"month": "พฤศจิกายน", "year": "2569"})
        template.save(str(raw_path))
        enforce_attendance_font_size(raw_path, final_path)
        return final_path

    def read_visible_run_sizes(self, path):
        """Map every visible run size (half-points) found in the package."""
        sizes = {}
        with ZipFile(path) as document:
            for item_name in document.namelist():
                if not item_name.startswith("word/") or not item_name.endswith(".xml"):
                    continue
                root = etree.fromstring(document.read(item_name))
                for run in root.iter(f"{NAMESPACE}r"):
                    text = "".join(
                        node.text or "" for node in run.iter(f"{NAMESPACE}t")
                    )
                    if not text.strip():
                        continue
                    size = run.find(f"{NAMESPACE}rPr/{NAMESPACE}sz")
                    value = size.get(f"{NAMESPACE}val") if size is not None else None
                    sizes[value] = sizes.get(value, 0) + 1
        return sizes

    def test_every_visible_run_is_sixteen_points(self):
        for category in TEMPLATES:
            with self.subTest(category=category), tempfile.TemporaryDirectory() as directory:
                final_path = self.render_attendance_sheet(category, directory)
                sizes = self.read_visible_run_sizes(final_path)

                self.assertTrue(sizes, "expected visible text in the attendance sheet")
                self.assertEqual(
                    set(sizes),
                    {"32"},
                    f"{category} sheet has runs that are not 16 pt: {sizes}",
                )

    def read_body_paragraphs(self, path):
        """Return the body paragraphs of a document, as XML elements."""
        with ZipFile(path) as document:
            root = etree.fromstring(document.read("word/document.xml"))
        body = root.find(f"{NAMESPACE}body")
        return body.findall(f"{NAMESPACE}p")

    def test_trailing_empty_paragraphs_are_removed(self):
        with tempfile.TemporaryDirectory() as directory:
            raw_path = Path(directory) / "raw.docx"
            final_path = Path(directory) / "final.docx"
            template_path = self.project_dir / TEMPLATES["guard"]["Table"]
            docx_template_path = template_path
            if template_path.suffix.lower() == ".doc":
                docx_template_path = Path(directory) / "converted.docx"
                convert_legacy_doc(template_path, docx_template_path)

            prepared_path = Path(directory) / "prepared.docx"
            prepare_template(docx_template_path, prepared_path)
            template = DocxTemplate(str(prepared_path))
            template.render({"month": "พฤศจิกายน", "year": "2569"})
            template.save(str(raw_path))
            enforce_attendance_font_size(raw_path, final_path)

            before = self.read_body_paragraphs(raw_path)
            after = self.read_body_paragraphs(final_path)

        self.assertLess(
            len(after),
            len(before),
            "the trailing empty paragraph of the guard sheet should be removed",
        )

        def shows_content(paragraph):
            text = "".join(
                node.text or "" for node in paragraph.iter(f"{NAMESPACE}t")
            )
            return bool(
                text.strip()
                or paragraph.find(f".//{NAMESPACE}drawing") is not None
                or paragraph.find(f".//{NAMESPACE}pict") is not None
                or paragraph.find(f".//{NAMESPACE}object") is not None
            )

        self.assertTrue(
            shows_content(after[-1]),
            "the document still ends with an empty paragraph",
        )
        self.assertEqual(
            len(after),
            len(before) - 1,
            "only the single trailing empty paragraph should be removed",
        )

    def test_guard_sheet_fits_on_a_single_page(self):
        if shutil.which(WORDCRAFT_CLI) is None:
            self.skipTest(f"WordCraft CLI not available: {WORDCRAFT_CLI}")

        with tempfile.TemporaryDirectory() as directory:
            final_path = self.render_attendance_sheet("guard", directory)
            pdf_path = Path(directory) / "final.pdf"
            convert_to_pdf(final_path, pdf_path)

            with pymupdf.open(pdf_path) as document:
                self.assertEqual(
                    document.page_count,
                    1,
                    "the guard attendance sheet must fit on one page",
                )
                sizes = {
                    round(span["size"], 1)
                    for block in document[0].get_text("dict")["blocks"]
                    if block["type"] == 0
                    for line in block["lines"]
                    for span in line["spans"]
                    if span["text"].strip()
                }

        self.assertEqual(sizes, {16.0}, f"guard sheet renders at {sorted(sizes)} pt")


class AttendanceDocumentTests(unittest.TestCase):
    def test_buddhist_year_requires_four_ascii_digits(self):
        self.assertTrue(is_valid_buddhist_year("2569"))
        self.assertFalse(is_valid_buddhist_year("256"))
        self.assertFalse(is_valid_buddhist_year("๒๕๖๙"))

    def test_combined_pdf_opens_with_cover_before_table(self):
        documents = {
            "Cover": {"pdf_bytes": self.create_pdf(("Cover page 1", "Cover page 2"))},
            "Table": {"pdf_bytes": self.create_pdf(("Table page 1", "Table page 2"))},
        }

        combined_bytes = combine_pdf_bytes(documents)
        with pymupdf.open(stream=combined_bytes, filetype="pdf") as combined:
            self.assertEqual(combined.page_count, 4)
            page_text = [page.get_text() for page in combined]

        self.assertIn("Cover page 1", page_text[0])
        self.assertIn("Cover page 2", page_text[1])
        self.assertIn("Table page 1", page_text[2])
        self.assertIn("Table page 2", page_text[3])

    def test_docx_templates_render_month_and_year_placeholders(self):
        project_dir = Path(__file__).resolve().parents[1]
        docx_templates = (
            project_dir / template_name
            for templates in TEMPLATES.values()
            for template_name in templates.values()
            if template_name.lower().endswith(".docx")
        )

        for template_path in docx_templates:
            with self.subTest(template=template_path.name), tempfile.TemporaryDirectory() as directory:
                prepared_path = Path(directory) / "prepared.docx"
                rendered_path = Path(directory) / "rendered.docx"
                prepare_template(template_path, prepared_path)
                template = DocxTemplate(str(prepared_path))
                template.render({"month": "พฤศจิกายน", "year": "2569"})
                template.save(str(rendered_path))

                text = self.read_docx_text(rendered_path)
                self.assertIn("พฤศจิกายน", text)
                self.assertIn("2569", text)
                self.assertNotIn("{{month}}", text)
                self.assertNotIn("{{year}}", text)

    def test_template_signature_detects_content_changes_with_same_size_and_mtime(self):
        project_dir = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            for template_name in TEMPLATES["temporary"].values():
                source = project_dir / template_name
                target = temp_dir / template_name
                target.write_bytes(source.read_bytes())
                target.touch()

            original_signature = get_template_signature(temp_dir, "temporary")
            modified_template = temp_dir / next(
                iter(TEMPLATES["temporary"].values())
            )
            original_stat = modified_template.stat()
            original_bytes = modified_template.read_bytes()
            modified_template.write_bytes(
                bytes((original_bytes[0] ^ 1,)) + original_bytes[1:]
            )
            os.utime(
                modified_template,
                ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns),
            )
            modified_size = modified_template.stat().st_size

            updated_signature = get_template_signature(temp_dir, "temporary")

        self.assertEqual(len(original_bytes), modified_size)
        self.assertNotEqual(original_signature, updated_signature)

    def test_preview_cache_reuses_files_and_prunes_old_or_excess_entries(self):
        with tempfile.TemporaryDirectory() as directory:
            static_dir = Path(directory)
            first = save_preview_pdfs(
                {"Combined": {"pdf_bytes": b"%PDF-first"}},
                static_dir,
                ("temporary", "พฤศจิกายน", 2569),
            )
            repeated = save_preview_pdfs(
                {"Combined": {"pdf_bytes": b"%PDF-first"}},
                static_dir,
                ("temporary", "พฤศจิกายน", 2569),
            )
            self.assertEqual(first, repeated)

            preview_dir = static_dir / "preview"
            now = 1_800_000_000
            protected_path = preview_dir / Path(first["Combined"]).name
            for index in range(PREVIEW_CACHE_MAX_FILES + 5):
                preview_path = preview_dir / f"preview_test_{index}.pdf"
                preview_path.write_bytes(b"%PDF-test")
                modified = now - index
                if index == PREVIEW_CACHE_MAX_FILES + 4:
                    modified = now - PREVIEW_CACHE_MAX_AGE_SECONDS - 1
                os.utime(preview_path, (modified, modified))

            prune_preview_pdfs(
                preview_dir,
                protected_paths={protected_path},
                now=now,
            )

            remaining_count = len(list(preview_dir.glob("preview_*.pdf")))
            self.assertLessEqual(remaining_count, PREVIEW_CACHE_MAX_FILES)
            self.assertTrue(protected_path.is_file())
            self.assertFalse(
                (
                    preview_dir
                    / f"preview_test_{PREVIEW_CACHE_MAX_FILES + 4}.pdf"
                ).exists()
            )


    @staticmethod
    def create_pdf(page_labels):
        with pymupdf.open() as document:
            for label in page_labels:
                page = document.new_page()
                page.insert_text((72, 72), label)
            return document.tobytes()

    @staticmethod
    def read_docx_text(path):
        namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        text_nodes = []
        with ZipFile(path) as document:
            for item_name in document.namelist():
                if item_name.startswith("word/") and item_name.endswith(".xml"):
                    root = etree.fromstring(document.read(item_name))
                    text_nodes.extend(
                        node.text or "" for node in root.iter(f"{namespace}t")
                    )
        return "".join(text_nodes)


if __name__ == "__main__":
    unittest.main()
