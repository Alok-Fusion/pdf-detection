import io
import re
import json
import zipfile

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd

import pytesseract
from PIL import Image


# --------- PAGE CONFIG ---------
st.set_page_config(layout="wide")


# --------- REGEX ---------
MARK_REGEX = re.compile(r'\b[A-Z]{1,4}-\d+\b', re.IGNORECASE)


# --------- TEXT UTILITIES ---------

def normalize_text(text: str) -> str:
    return text.replace("—", "-").replace("–", "-")


def extract_marks_from_text(text: str, marks_set: set):
    text = normalize_text(text)
    for m in MARK_REGEX.findall(text):
        marks_set.add(m.upper())


# ======================================================
# OCR LAYER (MAKES PDF FULLY SEARCHABLE)
# ======================================================

def ocr_pdf(pdf_bytes: bytes, dpi=250) -> bytes:
    """
    Memory-safe OCR for large mechanical / architectural PDFs.
    Uses Tesseract to generate searchable PDF pages directly.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()

    for page_index, page in enumerate(src):
        # Render page (lower DPI to avoid memory blowup)
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Let Tesseract generate a searchable PDF page
        pdf_bytes_ocr = pytesseract.image_to_pdf_or_hocr(
            img, extension="pdf"
        )

        ocr_page = fitz.open(stream=pdf_bytes_ocr, filetype="pdf")
        out.insert_pdf(ocr_page)

        # Explicit cleanup (important!)
        pix = None
        img = None
        ocr_page.close()

    buf = io.BytesIO()
    out.save(buf)
    out.close()
    src.close()

    buf.seek(0)
    return buf.getvalue()

# ======================================================
# CORE LOGIC (UNCHANGED)
# ======================================================

def extract_schedules_and_marks(pdf_bytes: bytes):
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = normalize_text(page.extract_text() or "")
            extract_marks_from_text(text, marks_set)

            for table in page.extract_tables() or []:
                for row in table:
                    for cell in row:
                        if cell:
                            extract_marks_from_text(str(cell), marks_set)

    return sorted(marks_set)


def mark_type(mark: str) -> str:
    m = re.match(r'^([A-Z]{1,4})', mark)
    return m.group(1).upper() if m else mark.split("-")[0].upper()


def get_plan_label(page: fitz.Page):
    lines = [l.strip() for l in (page.get_text() or "").splitlines() if l.strip()]
    plans = [l for l in lines if "PLAN" in l.upper()]
    return plans[0] if plans else None


def build_type_color_map(types):
    palette = [
        (1, 0, 0), (0, 0, 1), (0, 0.6, 0),
        (1, 0.5, 0), (0.6, 0, 0.6),
        (0, 0.7, 0.7), (0.7, 0.7, 0),
    ]
    return {t: palette[i % len(palette)] for i, t in enumerate(sorted(set(types)))}


def highlight_pdf_and_collect(pdf_bytes, marks, file_name):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    type_color_map = build_type_color_map([mark_type(m) for m in marks])

    rows = []

    for page_index, page in enumerate(doc):
        plan_label = get_plan_label(page)

        for mark in marks:
            m_type = mark_type(mark)
            color = type_color_map[m_type]

            variants = {
                mark,
                mark.replace("-", " "),
                mark.replace("-", ""),
            }

            rects = []
            for v in variants:
                rects += page.search_for(v)

            rects = list(set(rects))
            if not rects:
                continue

            for r in rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()

            rows.append({
                "file_name": file_name,
                "plan_label": plan_label,
                "page_number": page_index + 1,
                "mark": mark,
                "mark_type": m_type,
                "count_on_page": len(rects),
                "color_r": color[0],
                "color_g": color[1],
                "color_b": color[2],
            })

    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    buf.seek(0)

    return buf.getvalue(), pd.DataFrame(rows)


# ======================================================
# STREAMLIT UI
# ======================================================

st.title("Mechanical PDF → Searchable → CSV + JSON")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if "processed" not in st.session_state:
    st.session_state.processed = False


if uploaded_file and not st.session_state.processed:
    with st.spinner("🔍 Running OCR to make PDF searchable..."):
        original_pdf = uploaded_file.read()
        searchable_pdf = ocr_pdf(original_pdf)

    with st.spinner("📊 Extracting marks & highlighting..."):
        marks = extract_schedules_and_marks(searchable_pdf)

        if not marks:
            st.error("No marks detected even after OCR.")
            st.stop()

        highlighted_pdf, master_df = highlight_pdf_and_collect(
            searchable_pdf, marks, uploaded_file.name
        )

        master_json = {
            "file_name": uploaded_file.name,
            "records": [
                {
                    "plan_label": row.plan_label,
                    "page_number": int(row.page_number),
                    "mark": row.mark,
                    "mark_type": row.mark_type,
                    "count_on_page": int(row.count_on_page),
                    "color": {
                        "r": row.color_r,
                        "g": row.color_g,
                        "b": row.color_b,
                    },
                }
                for row in master_df.itertuples(index=False)
            ],
        }

        st.session_state.master_df = master_df
        st.session_state.master_json = master_json
        st.session_state.highlighted_pdf = highlighted_pdf
        st.session_state.original_pdf = original_pdf
        st.session_state.file_name = uploaded_file.name
        st.session_state.processed = True


# ======================================================
# OUTPUT
# ======================================================

if st.session_state.processed:
    st.subheader("📊 Extracted Data")
    st.dataframe(st.session_state.master_df, use_container_width=True, height=450)

    st.subheader("🧾 JSON Output")
    st.json(st.session_state.master_json)

    csv_bytes = st.session_state.master_df.to_csv(index=False).encode("utf-8")
    json_bytes = json.dumps(st.session_state.master_json, indent=2).encode("utf-8")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.writestr(f"input/{st.session_state.file_name}", st.session_state.original_pdf)
        zipf.writestr(
            f"output/{st.session_state.file_name.rsplit('.',1)[0]}_highlighted.pdf",
            st.session_state.highlighted_pdf
        )
        zipf.writestr("data/master_data.csv", csv_bytes)
        zipf.writestr("data/master_data.json", json_bytes)

    zip_buffer.seek(0)

    st.download_button(
        "⬇️ Download ZIP (PDF + CSV + JSON)",
        data=zip_buffer.getvalue(),
        file_name=f"{st.session_state.file_name.rsplit('.',1)[0]}_results.zip",
        mime="application/zip",
    )
