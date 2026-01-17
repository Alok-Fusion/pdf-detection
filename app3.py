import io
import re
import json
from collections import defaultdict

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image
import pytesseract


# ===================== CONFIG =====================

st.set_page_config(
    page_title="Mechanical Schedules → Mark Highlighter",
    page_icon="📐",
    layout="wide",
)

# More tolerant regex for OCR + native PDFs
MARK_REGEX = re.compile(r'[A-Z]{1,4}\s*-\s*\d+', re.IGNORECASE)


# ===================== PDF READABILITY =====================

def ensure_searchable_pdf(pdf_bytes: bytes) -> bytes:
    """
    If PDF text is extractable → return original.
    If not → OCR and return searchable PDF.
    """
    # Check first few pages for text
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:3]:
            text = page.extract_text()
            if text and len(text.strip()) > 50:
                return pdf_bytes  # already readable

    # OCR fallback
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    new_doc = fitz.open()

    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))

        ocr_text = pytesseract.image_to_string(img)

        new_page = new_doc.new_page(
            width=page.rect.width,
            height=page.rect.height
        )

        # Insert original image
        new_page.insert_image(page.rect, stream=pix.tobytes("png"))

        # Insert invisible OCR text layer
        new_page.insert_textbox(
            page.rect,
            ocr_text,
            fontsize=6,
            overlay=True
        )

    out = io.BytesIO()
    new_doc.save(out)
    new_doc.close()
    doc.close()
    out.seek(0)

    return out.getvalue()


# ===================== EXTRACTION =====================

def extract_schedules_and_marks(pdf_bytes: bytes):
    schedule_tables = []
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if "SCHEDULE" not in text.upper():
                continue

            tables = page.extract_tables()
            for t_index, table in enumerate(tables or []):
                if not table:
                    continue

                cleaned = [
                    [("" if c is None else str(c).strip()) for c in row]
                    for row in table
                ]

                if all(all(cell == "" for cell in row) for row in cleaned):
                    continue

                header_idx = next(
                    (i for i, r in enumerate(cleaned) if any(r)), None
                )
                if header_idx is None:
                    continue

                header = cleaned[header_idx]
                data_rows = [
                    r for r in cleaned[header_idx + 1:]
                    if any(r)
                ]

                header_safe = [
                    col if col else f"COL_{i+1}"
                    for i, col in enumerate(header)
                ]

                dict_rows = []
                for r in data_rows:
                    row_dict = {}
                    for i, col in enumerate(header_safe):
                        row_dict[col] = r[i] if i < len(r) else ""
                    dict_rows.append(row_dict)

                schedule_tables.append({
                    "page_index": page_index,
                    "table_index_on_page": t_index,
                    "header": header_safe,
                    "rows": dict_rows
                })

                # Extract marks from cells
                for r in data_rows:
                    for cell in r:
                        if not cell:
                            continue
                        m = MARK_REGEX.search(cell)
                        if m:
                            marks_set.add(
                                m.group(0).upper().replace(" ", "")
                            )

    return {
        "schedule_tables": schedule_tables,
        "marks": sorted(marks_set),
    }, sorted(marks_set)


# ===================== HIGHLIGHTING =====================

def mark_type(mark: str) -> str:
    return mark.split("-")[0]


def get_plan_label(page: fitz.Page):
    lines = [l.strip() for l in page.get_text().splitlines() if l.strip()]
    plans = [l for l in lines if "PLAN" in l.upper()]
    return plans[0] if plans else None


def build_type_color_map(types):
    palette = [
        (1, 0, 0), (0, 0, 1), (0, 0.6, 0),
        (1, 0.5, 0), (0.6, 0, 0.6),
        (0, 0.7, 0.7), (0.5, 0.5, 0.5)
    ]
    return {
        t: palette[i % len(palette)]
        for i, t in enumerate(sorted(set(types)))
    }


def highlight_pdf(pdf_bytes: bytes, marks: list[str]):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    types = [mark_type(m) for m in marks]
    type_color_map = build_type_color_map(types)

    plan_mark_counts = defaultdict(lambda: defaultdict(int))
    plan_type_counts = defaultdict(lambda: defaultdict(int))

    for page in doc:
        plan_label = get_plan_label(page)

        for mark in marks:
            t = mark_type(mark)
            color = type_color_map[t]

            variants = {
                mark,
                mark.replace("-", " "),
                mark.replace("-", ""),
            }

            rects = []
            for v in variants:
                rects += page.search_for(v)

            unique_rects = list({tuple(r): r for r in rects}.values())

            for r in unique_rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()

            if plan_label:
                plan_mark_counts[plan_label][mark] += len(unique_rects)
                plan_type_counts[plan_label][t] += len(unique_rects)

    out = io.BytesIO()
    doc.save(out)
    doc.close()
    out.seek(0)

    return (
        out.getvalue(),
        dict(plan_mark_counts),
        dict(plan_type_counts),
        type_color_map
    )


# ===================== STREAMLIT UI =====================

st.title("📐 Mechanical Schedules → Mark Highlighter")

st.markdown(
    """
    **Pipeline**
    1. Make PDF readable (OCR if needed)
    2. Extract marks from schedules
    3. Highlight marks on all sheets
    """
)

uploaded_file = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

if uploaded_file:
    raw_pdf_bytes = uploaded_file.read()

    with st.spinner("Checking PDF readability..."):
        pdf_bytes = ensure_searchable_pdf(raw_pdf_bytes)

    with st.spinner("Extracting schedules and marks..."):
        schedule_json, marks = extract_schedules_and_marks(pdf_bytes)

    st.subheader("Detected Marks")
    if marks:
        st.code(", ".join(marks))
    else:
        st.warning("No marks detected.")

    with st.expander("Schedule JSON (preview)"):
        st.json(schedule_json)

    if marks:
        with st.spinner("Highlighting marks..."):
            highlighted_pdf, plan_mark_counts, plan_type_counts, type_color_map = (
                highlight_pdf(pdf_bytes, marks)
            )

        st.download_button(
            "⬇️ Download Highlighted PDF",
            highlighted_pdf,
            "highlighted.pdf",
            "application/pdf"
        )

        st.subheader("Counts per Plan (by mark)")
        st.json(plan_mark_counts)

        st.subheader("Counts per Plan (by type)")
        st.json(plan_type_counts)

        st.subheader("Type → Color Map")
        st.json(type_color_map)
