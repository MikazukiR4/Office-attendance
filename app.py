"""Streamlit app for generating monthly attendance covers and tables."""

import hashlib
import os
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

import streamlit as st

from attendance_config import (
    BUDDHIST_YEAR_MAX,
    BUDDHIST_YEAR_MIN,
    MONTHS,
    TEMPLATES,
    is_valid_buddhist_year,
)


CATEGORY_LABELS = {
    "temporary": "ลูกจ้างชั่วคราว",
    "government": "พนักงานราชการ",
    "guard": "ยามรักษาการณ์",
}
PREVIEW_CACHE_MAX_FILES = 36
PREVIEW_CACHE_MAX_AGE_SECONDS = 30 * 24 * 60 * 60


def get_resource_dir():
    """Return the bundled resource directory or the source directory."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def get_output_dir():
    """Keep generated documents beside a packaged app or source project."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "output"
    return Path(__file__).resolve().parent / "output"


@st.cache_data(show_spinner=False, max_entries=12)
def render_document_bytes(project_dir, category, month, year, template_signature):
    """Render templates and cache their PDF bytes until a template changes."""
    from docxtpl import DocxTemplate

    from generate_attendance import (
        convert_legacy_doc,
        initialized_com,
        prepare_template,
    )

    from docx2pdf import convert

    project_dir = Path(project_dir)
    generated = {}

    with tempfile.TemporaryDirectory() as temporary_dir:
        temporary_dir = Path(temporary_dir)
        rendered_dir = temporary_dir / "rendered"
        pdf_dir = temporary_dir / "pdf"
        rendered_dir.mkdir()
        pdf_dir.mkdir()

        for document_type, template_name in TEMPLATES[category].items():
            template_path = project_dir / template_name
            if not template_path.is_file():
                raise FileNotFoundError(f"Template not found: {template_path}")

            docx_template_path = template_path
            if template_path.suffix.lower() == ".doc":
                docx_template_path = temporary_dir / f"{template_path.stem}.docx"
                convert_legacy_doc(template_path, docx_template_path)

            prepared_path = temporary_dir / f"prepared_{document_type}.docx"
            prepare_template(docx_template_path, prepared_path)

            docx_path = rendered_dir / f"{document_type}.docx"

            template = DocxTemplate(str(prepared_path))
            template.render({"month": month, "year": str(year)})
            template.save(str(docx_path))

        with initialized_com():
            convert(str(rendered_dir), str(pdf_dir))

        for document_type in TEMPLATES[category]:
            docx_path = rendered_dir / f"{document_type}.docx"
            pdf_path = pdf_dir / f"{document_type}.pdf"
            generated[document_type] = {
                "docx_bytes": docx_path.read_bytes(),
                "pdf_bytes": pdf_path.read_bytes(),
            }

        generated["Combined"] = {
            "pdf_bytes": combine_pdf_bytes(generated)
        }

    return generated


def combine_pdf_bytes(documents):
    """Combine cover and attendance PDFs in the required order."""
    import pymupdf

    with pymupdf.open() as combined_pdf:
        for document_type in ("Cover", "Table"):
            document_pdf = pymupdf.open(
                stream=documents[document_type]["pdf_bytes"],
                filetype="pdf",
            )
            try:
                combined_pdf.insert_pdf(document_pdf)
            finally:
                document_pdf.close()
        return combined_pdf.tobytes()


def make_print_pdf(documents, document_type, copies):
    """Create a PDF containing one selected document repeated for printing."""
    import pymupdf

    if copies < 1:
        raise ValueError("copies must be at least 1")

    with pymupdf.open() as print_pdf:
        source_pdf = pymupdf.open(
            stream=documents[document_type]["pdf_bytes"],
            filetype="pdf",
        )
        try:
            for _ in range(copies):
                print_pdf.insert_pdf(source_pdf)
        finally:
            source_pdf.close()
        return print_pdf.tobytes()


def generate_documents(project_dir, output_dir, category, month, year):
    """Write generated Word files and the combined PDF to the output folder."""
    project_dir = Path(project_dir)
    output_dir = Path(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    template_signature = get_template_signature(project_dir, category)
    generated = render_document_bytes(
        str(project_dir),
        category,
        month,
        int(year),
        template_signature,
    )

    for document_type, template_name in TEMPLATES[category].items():
        docx_path = output_dir / (
            f"{category}_{document_type}_{month}_{year}.docx"
        )
        docx_path.write_bytes(generated[document_type]["docx_bytes"])
        pdf_name = f"{category}_{document_type}_{month}_{year}.pdf"
        generated[document_type]["file_name"] = pdf_name

    combined_path = output_dir / f"{category}_{month}_{year}.pdf"
    combined_path.write_bytes(generated["Combined"]["pdf_bytes"])
    generated["Combined"]["file_name"] = combined_path.name

    return generated


def get_template_signature(project_dir, category):
    """Validate templates and return a cache key based on their contents."""
    for template_name in TEMPLATES[category].values():
        template_path = project_dir / template_name
        if not template_path.is_file():
            raise FileNotFoundError(f"Template not found: {template_path}")
    return tuple(
        (
            template_name,
            hashlib.sha256((project_dir / template_name).read_bytes()).hexdigest(),
        )
        for template_name in TEMPLATES[category].values()
    )


def save_preview_pdfs(documents, static_dir, cache_key):
    """Write cached preview PDFs and prune expired or excess files."""
    preview_dir = static_dir / "preview"
    preview_dir.mkdir(parents=True, exist_ok=True)

    token = hashlib.sha256(repr(cache_key).encode("utf-8")).hexdigest()[:20]
    urls = {}
    protected_paths = set()
    for document_type, document in documents.items():
        filename = f"preview_{token}_{document_type.lower()}.pdf"
        preview_path = preview_dir / filename
        if not preview_path.is_file():
            preview_path.write_bytes(document["pdf_bytes"])
        protected_paths.add(preview_path)
        urls[document_type] = f"/app/static/preview/{filename}"
    prune_preview_pdfs(preview_dir, protected_paths)
    return urls


def prune_preview_pdfs(preview_dir, protected_paths=(), now=None):
    """Keep active previews and prune old files beyond the configured cache limit."""
    now = now if now is not None else time.time()
    expiry = now - PREVIEW_CACHE_MAX_AGE_SECONDS
    protected_paths = {path.resolve() for path in protected_paths}
    files = sorted(
        preview_dir.glob("preview_*.pdf"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    kept_count = sum(path.resolve() in protected_paths for path in files)
    for preview_path in files:
        if preview_path.resolve() in protected_paths:
            continue
        elif (
            preview_path.stat().st_mtime < expiry
            or kept_count >= PREVIEW_CACHE_MAX_FILES
        ):
            preview_path.unlink()
        else:
            kept_count += 1


def main():
    st.set_page_config(
        page_title="ระบบสร้างเอกสารลงเวลา",
        page_icon="🗓️",
        layout="wide",
    )
    header_column, theme_column = st.columns((5, 1), vertical_alignment="center")
    if not st.session_state.get("_dark_mode_default_set"):
        st.session_state["dark_mode"] = True
        st.session_state["_dark_mode_default_set"] = True

    with theme_column:
        dark_mode = st.toggle("โหมดมืด", key="dark_mode")

    if dark_mode:
        palette = {
            "background": "#0b1220",
            "surface": "#111c2e",
            "surface_alt": "#17243a",
            "border": "#29384d",
            "text": "#e5edf7",
            "muted": "#a8b6c9",
            "accent": "#34b8a5",
            "accent_hover": "#2aa18f",
            "glow": "rgba(20, 105, 95, .20)",
            "field_text": "#e5edf7",
            "button_text": "#e5edf7",
            "button_hover": "#22334b",
        }
    else:
        palette = {
            "background": "#f5f7fa",
            "surface": "rgba(255, 255, 255, .92)",
            "surface_alt": "#ffffff",
            "border": "#e4eaf0",
            "text": "#1e293b",
            "muted": "#64748b",
            "accent": "#117c70",
            "accent_hover": "#0d685e",
            "glow": "rgba(231, 244, 241, .9)",
            "field_text": "#1e293b",
            "button_text": "#1e293b",
            "button_hover": "#f1f5f9",
        }

    st.markdown(
        f"""
        <style>
        :root {{
            --app-background: {palette["background"]};
            --app-surface: {palette["surface"]};
            --app-surface-alt: {palette["surface_alt"]};
            --app-border: {palette["border"]};
            --app-text: {palette["text"]};
            --app-muted: {palette["muted"]};
            --app-accent: {palette["accent"]};
            --app-accent-hover: {palette["accent_hover"]};
            --app-glow: {palette["glow"]};
            --app-field-text: {palette["field_text"]};
            --app-button-text: {palette["button_text"]};
            --app-button-hover: {palette["button_hover"]};
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(ellipse at 12% 0%, var(--app-glow) 0, transparent 34%),
                var(--app-background);
            color: var(--app-text) !important;
        }
        [data-testid="stMainBlockContainer"] {
            max-width: 1600px;
            padding-top: 2.5rem;
            padding-bottom: 3rem;
            padding-left: clamp(1rem, 3vw, 3.5rem);
            padding-right: clamp(1rem, 3vw, 3.5rem);
        }
        h1, h2, h3, h4, p, label,
        [data-testid="stCaptionContainer"],
        [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"] {
            color: var(--app-text) !important;
        }
        h1, h2, h3, p, label, button, input {
            font-family: "Leelawadee UI", "Tahoma", sans-serif;
        }
        [data-testid="stSelectbox"] [role="group"] {
            background: var(--app-surface-alt) !important;
            border-color: var(--app-border) !important;
        }
        [data-testid="stSelectbox"] input {
            color: var(--app-field-text) !important;
        }
        [data-testid="stSelectbox"] svg {
            color: var(--app-muted) !important;
        }
        [data-testid="stNumberInput"] input {
            background: var(--app-surface-alt) !important;
            color: var(--app-field-text) !important;
            border-color: var(--app-border) !important;
        }
        .app-eyebrow {
            color: var(--app-accent);
            font-size: .9rem;
            font-weight: 700;
            letter-spacing: .04em;
            margin-bottom: .35rem;
        }
        .app-subtitle {
            color: var(--app-muted);
            margin-top: -.5rem;
            margin-bottom: 1.7rem;
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--app-surface);
            border: 1px solid var(--app-border);
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, .12);
        }
        .stHorizontalBlock:has(iframe) > [data-testid="stColumn"]:first-child > [data-testid="stVerticalBlock"] {
            height: 100%;
            display: flex;
            flex-direction: column;
        }
        .stHorizontalBlock:has(iframe) > [data-testid="stColumn"]:first-child > [data-testid="stVerticalBlock"] > [data-testid="stLayoutWrapper"] {
            flex: 1 1 0;
        }
        [data-testid="stRadio"] [role="radiogroup"] {
            gap: .65rem;
        }
        [data-testid="stRadio"] label {
            border: 1px solid var(--app-border);
            border-radius: 12px;
            padding: .7rem 1rem;
            background: var(--app-surface-alt);
            transition: border-color .15s ease, background .15s ease;
        }
        [data-testid="stRadio"] label:has(input:checked) {
            border-color: var(--app-accent);
            background: var(--app-surface);
        }
        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {
            min-height: 2.8rem;
            border-radius: 10px;
            font-weight: 650;
        }
        [data-testid="stButton"] button[kind="primary"],
        [data-testid="stDownloadButton"] button[kind="primary"] {
            background: var(--app-accent);
            border-color: var(--app-accent);
            color: #fff !important;
        }
        [data-testid="stButton"] button[kind="primary"]:hover,
        [data-testid="stDownloadButton"] button[kind="primary"]:hover {
            background: var(--app-accent-hover);
            border-color: var(--app-accent-hover);
            color: #fff !important;
        }
        [data-testid="stButton"] button:not([kind="primary"]),
        [data-testid="stDownloadButton"] button:not([kind="primary"]) {
            background: var(--app-surface-alt);
            border: 1px solid var(--app-border);
            color: var(--app-button-text) !important;
        }
        [data-testid="stButton"] button:not([kind="primary"]):hover,
        [data-testid="stDownloadButton"] button:not([kind="primary"]):hover {
            background: var(--app-button-hover);
            border-color: var(--app-muted);
            color: var(--app-text) !important;
        }
        [data-testid="stButton"] button[kind="primary"]:disabled {
            background: var(--app-muted);
            border-color: var(--app-muted);
            color: #fff !important;
        }
        [data-testid="stAlert"] p,
        [data-testid="stAlert"] div {
            color: var(--app-text) !important;
        }
        [data-baseweb="popover"],
        [role="listbox"],
        [role="option"] {
            background: var(--app-surface-alt) !important;
            color: var(--app-text) !important;
        }
        @media (max-width: 900px) {
            [data-testid="stMainBlockContainer"] {
                padding-top: 1.5rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    today = date.today()
    next_month_number = today.month % 12 + 1
    next_month_year = today.year + int(today.month == 12) + 543
    if "attendance_month" not in st.session_state:
        st.session_state["attendance_month"] = MONTHS[next_month_number - 1]
    if "attendance_year" not in st.session_state:
        st.session_state["attendance_year"] = next_month_year

    with header_column:
        st.markdown('<div class="app-eyebrow">สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม</div>', unsafe_allow_html=True)
        st.title("ระบบสร้างเอกสารลงเวลา")
        st.markdown(
            '<p class="app-subtitle">สร้างปกและบัญชีลงเวลารายเดือน พร้อมดูตัวอย่างและดาวน์โหลด PDF</p>',
            unsafe_allow_html=True,
        )

    form_column, preview_column = st.columns((1, 1), gap="large")
    with form_column:
        with st.container(border=True):
            st.subheader("1 · เลือกประเภทบุคลากร")
            selected_category = st.radio(
                "ประเภทบุคลากร",
                options=tuple(CATEGORY_LABELS),
                format_func=lambda key: CATEGORY_LABELS[key],
                horizontal=True,
                label_visibility="collapsed",
            )

        with st.container(border=True):
            st.subheader("2 · เลือกเดือนและปี")
            month_column, year_column = st.columns((1.4, 1))
            with month_column:
                month = st.selectbox("เดือนที่จัดทำ", MONTHS, key="attendance_month")
            with year_column:
                year = st.number_input(
                    "ปี พ.ศ.",
                    min_value=BUDDHIST_YEAR_MIN,
                    max_value=BUDDHIST_YEAR_MAX,
                    step=1,
                    key="attendance_year",
                )
            if not is_valid_buddhist_year(year):
                st.error("กรุณาระบุปี พ.ศ. เป็นตัวเลข 4 หลัก")
                st.stop()
            st.caption("ตั้งต้นเป็นเดือนถัดไปจากปัจจุบัน สามารถเปลี่ยนเดือนหรือปีได้")

        with st.container(border=True):
            st.subheader("3 · ดำเนินการ")
            print_page = st.selectbox(
                "หน้าที่ต้องการพิมพ์",
                options=("Cover", "Table"),
                format_func=lambda value: {
                    "Cover": "หน้า ปก",
                    "Table": "หน้า ตารางลงเวลา",
                }[value],
            )
            print_copies = st.number_input(
                "จำนวนแผ่น",
                min_value=1,
                max_value=100,
                value=1,
                step=1,
            )
            generate_clicked = st.button(
                "สร้าง PDF ฉบับรวม",
                type="primary",
                use_container_width=True,
            )

        project_dir = get_resource_dir()
        preview_key = None
        preview_url = None
        preview_cache = st.session_state.setdefault("document_preview_cache", {})
        try:
            template_signature = get_template_signature(
                project_dir,
                selected_category,
            )
            preview_key = (
                selected_category,
                month,
                int(year),
                template_signature,
            )
            preview_url = preview_cache.get(preview_key)
            preview_file_exists = (
                preview_url
                and (
                    project_dir / "static" / "preview" / Path(preview_url).name
                ).is_file()
            )
            if not preview_file_exists:
                preview_url = None
                with st.spinner("กำลังเตรียมตัวอย่าง PDF..."):
                    preview_documents = render_document_bytes(
                        str(project_dir),
                        selected_category,
                        month,
                        int(year),
                        template_signature,
                    )
                    preview_urls = save_preview_pdfs(
                        {"Combined": preview_documents["Combined"]},
                        project_dir / "static",
                        preview_key,
                    )
                    preview_url = preview_urls["Combined"]
            if preview_key in preview_cache:
                preview_cache.pop(preview_key)
            preview_cache[preview_key] = preview_url
            while len(preview_cache) > PREVIEW_CACHE_MAX_FILES:
                preview_cache.pop(next(iter(preview_cache)))
        except Exception as error:
            st.error(f"แสดงตัวอย่างเอกสารไม่สำเร็จ: {error}")
            st.exception(error)

        if generate_clicked:
            project_dir = get_resource_dir()
            output_dir = get_output_dir()
            try:
                generated = generate_documents(
                    project_dir,
                    output_dir,
                    selected_category,
                    month,
                    year,
                )
            except Exception as error:
                st.session_state.pop("generated_documents", None)
                st.error(f"สร้างเอกสารไม่สำเร็จ: {error}")
                st.exception(error)
            else:
                st.session_state["generated_documents"] = {
                    "category": selected_category,
                    "month": month,
                    "year": int(year),
                    "documents": generated,
                }
                st.success("สร้าง PDF ฉบับรวมเรียบร้อยแล้ว")

        result = st.session_state.get("generated_documents")
        if (
            result
            and result["category"] == selected_category
            and result["month"] == month
            and result["year"] == int(year)
        ):
            combined = result["documents"]["Combined"]
            st.download_button(
                "ดาวน์โหลด PDF ฉบับรวม",
                data=combined["pdf_bytes"],
                file_name=combined["file_name"],
                mime="application/pdf",
                key="download_combined",
                type="primary",
                use_container_width=True,
            )
            print_pdf = make_print_pdf(
                result["documents"],
                print_page,
                int(print_copies),
            )
            st.download_button(
                f"ดาวน์โหลด {print_page} สำหรับพิมพ์ {int(print_copies)} แผ่น",
                data=print_pdf,
                file_name=(
                    f"{selected_category}_{print_page}_{month}_{int(year)}_"
                    f"{int(print_copies)}copies.pdf"
                ),
                mime="application/pdf",
                key="download_print_page",
                use_container_width=True,
            )

    with preview_column:
        result = st.session_state.get("generated_documents")
        if preview_url:
            st.iframe(
                f"/app/static/embedpdf.html?src={preview_url}",
                height=900,
            )
        elif (
            result
            and result["category"] == selected_category
            and result["month"] == month
            and result["year"] == int(year)
        ):
            st.subheader("สร้างเอกสารเรียบร้อย")
            st.caption("PDF ฉบับรวมพร้อมดาวน์โหลดอยู่ทางด้านซ้าย")
            st.info("ลำดับเอกสาร: ปก → บัญชีลงเวลา")
        else:
            with st.container(border=True):
                st.subheader("ตัวอย่างเอกสาร")
                st.markdown("### ดู PDF ได้ในหน้านี้")
                st.write("เลือกประเภทบุคลากรหรือเปลี่ยนรอบเดือนทางซ้าย เพื่อดูตัวอย่าง PDF")
                st.caption("ตัวอย่างจะแสดงเป็นไฟล์เดียว เรียงจากปกไปยังบัญชีลงเวลา")


if __name__ == "__main__":
    main()
