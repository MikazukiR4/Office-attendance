"""Streamlit app for generating monthly attendance covers and tables."""

import os
import tempfile
import time
import uuid
from datetime import date
from pathlib import Path

import streamlit as st

from attendance_config import MONTHS, TEMPLATES


CATEGORY_LABELS = {
    "temporary": "ลูกจ้างชั่วคราว",
    "government": "พนักงานราชการ",
    "guard": "ยามรักษาการณ์",
}


@st.cache_data(show_spinner=False, max_entries=12)
def render_document_bytes(project_dir, category, month, year, template_signature):
    """Render templates and cache their PDF bytes until a template changes."""
    import pymupdf
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

        with pymupdf.open() as combined_pdf:
            for document_type in ("Cover", "Table"):
                document_pdf = pymupdf.open(
                    stream=generated[document_type]["pdf_bytes"],
                    filetype="pdf",
                )
                try:
                    combined_pdf.insert_pdf(document_pdf)
                finally:
                    document_pdf.close()
            generated["Combined"] = {"pdf_bytes": combined_pdf.tobytes()}

    return generated


def generate_documents(project_dir, output_dir, category, month, year):
    """Write generated Word files and the combined PDF to the output folder."""
    project_dir = Path(project_dir)
    output_dir = Path(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    for template_name in TEMPLATES[category].values():
        template_path = project_dir / template_name
        if not template_path.is_file():
            raise FileNotFoundError(f"Template not found: {template_path}")
    template_signature = tuple(
        (
            template_name,
            (project_dir / template_name).stat().st_mtime_ns,
            (project_dir / template_name).stat().st_size,
        )
        for template_name in TEMPLATES[category].values()
    )
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


def save_preview_pdfs(documents, static_dir):
    """Write preview PDFs to unique static URLs for the embedded viewer."""
    preview_dir = static_dir / "preview"
    preview_dir.mkdir(parents=True, exist_ok=True)

    expiry = time.time() - 24 * 60 * 60
    for stale_file in preview_dir.glob("preview_*.pdf"):
        if stale_file.stat().st_mtime < expiry:
            stale_file.unlink()

    token = uuid.uuid4().hex
    urls = {}
    for document_type, document in documents.items():
        filename = f"preview_{token}_{document_type.lower()}.pdf"
        (preview_dir / filename).write_bytes(document["pdf_bytes"])
        urls[document_type] = f"/app/static/preview/{filename}"
    return urls


def main():
    st.set_page_config(
        page_title="ระบบสร้างเอกสารลงเวลา",
        page_icon="🗓️",
        layout="wide",
    )
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(ellipse at 12% 0%, #e7f4f1 0, transparent 34%),
                #f5f7fa;
            color: #1e293b !important;
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
            color: #1e293b !important;
        }
        h1, h2, h3, p, label, button, input {
            font-family: "Leelawadee UI", "Tahoma", sans-serif;
        }
        [data-testid="stSelectbox"] [role="group"] {
            background: #fff !important;
            border-color: #cbd5e1 !important;
        }
        [data-testid="stSelectbox"] input {
            color: #1e293b !important;
        }
        [data-testid="stSelectbox"] svg {
            color: #475569 !important;
        }
        [data-testid="stNumberInput"] input {
            background: #fff !important;
            color: #1e293b !important;
        }
        .app-eyebrow {
            color: #117c70;
            font-size: .9rem;
            font-weight: 700;
            letter-spacing: .04em;
            margin-bottom: .35rem;
        }
        .app-subtitle {
            color: #64748b;
            margin-top: -.5rem;
            margin-bottom: 1.7rem;
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255, 255, 255, .92);
            border: 1px solid #e4eaf0;
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(15, 34, 52, .045);
        }
        [data-testid="stRadio"] [role="radiogroup"] {
            gap: .65rem;
        }
        [data-testid="stRadio"] label {
            border: 1px solid #dce5eb;
            border-radius: 12px;
            padding: .7rem 1rem;
            background: #fff;
            transition: border-color .15s ease, background .15s ease;
        }
        [data-testid="stRadio"] label:has(input:checked) {
            border-color: #16877a;
            background: #eff9f7;
        }
        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {
            min-height: 2.8rem;
            border-radius: 10px;
            font-weight: 650;
        }
        [data-testid="stButton"] button[kind="primary"],
        [data-testid="stDownloadButton"] button[kind="primary"] {
            background: #117c70;
            border-color: #117c70;
            color: #fff !important;
        }
        [data-testid="stButton"] button[kind="primary"]:hover {
            background: #0d685e;
            border-color: #0d685e;
            color: #fff !important;
        }
        [data-testid="stButton"] button:not([kind="primary"]),
        [data-testid="stDownloadButton"] button:not([kind="primary"]) {
            background: #fff;
            border: 1px solid #cbd5e1;
            color: #1e293b !important;
        }
        [data-testid="stButton"] button:not([kind="primary"]):hover,
        [data-testid="stDownloadButton"] button:not([kind="primary"]):hover {
            background: #f1f5f9;
            border-color: #94a3b8;
            color: #0f172a !important;
        }
        [data-testid="stButton"] button[kind="primary"]:disabled {
            background: #94a3b8;
            border-color: #94a3b8;
            color: #fff !important;
        }
        [data-testid="stAlert"] p,
        [data-testid="stAlert"] div {
            color: #1e293b !important;
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

    st.markdown('<div class="app-eyebrow">สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม</div>', unsafe_allow_html=True)
    st.title("ระบบสร้างเอกสารลงเวลา")
    st.markdown(
        '<p class="app-subtitle">สร้างปกและบัญชีลงเวลารายเดือน พร้อมดูตัวอย่างและดาวน์โหลด PDF</p>',
        unsafe_allow_html=True,
    )

    form_column, preview_column = st.columns((0.9, 1.1), gap="large")
    with form_column:
        with st.container(border=True):
            st.subheader("1 · เลือกประเภทบุคลากร")
            selected_category = st.selectbox(
                "ประเภทบุคลากร",
                options=tuple(CATEGORY_LABELS),
                format_func=lambda key: CATEGORY_LABELS[key],
            )

        with st.container(border=True):
            st.subheader("2 · เลือกเดือนและปี")
            month_column, year_column = st.columns((1.4, 1))
            with month_column:
                month = st.selectbox("เดือนที่จัดทำ", MONTHS, key="attendance_month")
            with year_column:
                year = st.number_input(
                    "ปี พ.ศ.",
                    min_value=1,
                    max_value=9999,
                    step=1,
                    key="attendance_year",
                )
            st.caption("ตั้งต้นเป็นเดือนถัดไปจากปัจจุบัน สามารถเปลี่ยนเดือนหรือปีได้")

        with st.container(border=True):
            st.subheader("3 · ดำเนินการ")
            preview_clicked = st.button(
                "ดูตัวอย่างก่อนสร้าง",
                use_container_width=True,
            )
            generate_clicked = st.button(
                "สร้าง PDF ฉบับรวม",
                type="primary",
                use_container_width=True,
            )

        if preview_clicked:
            project_dir = Path(__file__).resolve().parent
            try:
                with tempfile.TemporaryDirectory() as preview_dir:
                    preview_output_dir = Path(preview_dir) / "output"
                    preview_documents = generate_documents(
                        project_dir,
                        preview_output_dir,
                        selected_category,
                        month,
                        year,
                    )
                preview_urls = save_preview_pdfs(
                    {"Combined": preview_documents["Combined"]},
                    project_dir / "static",
                )
                st.session_state["document_preview"] = {
                    "category": selected_category,
                    "month": month,
                    "year": int(year),
                    "urls": preview_urls,
                }
            except Exception as error:
                st.session_state.pop("document_preview", None)
                st.error(f"แสดงตัวอย่างเอกสารไม่สำเร็จ: {error}")
                st.exception(error)

        if generate_clicked:
            project_dir = Path(__file__).resolve().parent
            output_dir = project_dir / "output"
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

    with preview_column:
        preview = st.session_state.get("document_preview")
        result = st.session_state.get("generated_documents")
        if (
            preview
            and preview["category"] == selected_category
            and preview["month"] == month
            and preview["year"] == int(year)
        ):
            preview_url = preview["urls"]["Combined"]
            st.subheader("ตัวอย่าง PDF ฉบับรวม")
            st.caption(f"{CATEGORY_LABELS[selected_category]} · {month} {year} · ปก → บัญชีลงเวลา")
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
                st.write("เลือกประเภทบุคลากรและรอบเดือนทางซ้าย จากนั้นกด “ดูตัวอย่างก่อนสร้าง”")
                st.caption("ตัวอย่างจะแสดงเป็นไฟล์เดียว เรียงจากปกไปยังบัญชีลงเวลา")


if __name__ == "__main__":
    main()
