import io
import re
import json
import os
import tempfile
import subprocess
from collections import defaultdict

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF


# ---------------- OCR UTILITY ----------------

def ensure_searchable_pdf(pdf_bytes: bytes) -> bytes:
    """
    Ensure the PDF has a searchable text layer.
    If searchable → return original bytes.
    If not → OCR using OCRmyPDF and return new bytes.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages[:3]:
                if page.extract_text():
                    return pdf_bytes
    except Exception:
        pass

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as in_pdf:
        in_pdf.write(pdf_bytes)
        in_path = in_pdf.name

    out_path = in_path.replace(".pdf", "_ocr.pdf")

    subprocess.run(
        [
            "ocrmypdf",
            "--force-ocr",
            "--deskew",
            "--rotate-pages",
            "--optimize", "3",
            in_path,
            out_path,
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    with open(out_path, "rb") as f:
        ocr_bytes = f.read()

    os.remove(in_path)
    os.remove(out_path)

    return ocr_bytes


# ---------------- UTILITIES ----------------

MARK_REGEX = re.compile(
    r'\b[A-Z]{1,5}-\d{1,3}\b',
    re.IGNORECASE
)


def extract_schedules_and_marks(pdf_bytes: bytes):
    schedule_tables = []
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = (page.extract_text() or "")
            text_upper = text.upper()

            # Only schedule pages
            if "SCHEDULE" not in text_upper:
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

                header_idx = next((i for i, r in enumerate(cleaned) if any(r)), None)
                if header_idx is None:
                    continue

                header = cleaned[header_idx]
                data_rows = [r for r in cleaned[header_idx + 1:] if any(r)]
                if not data_rows:
                    continue

                header_safe = [
                    h if h else f"COL_{i+1}" for i, h in enumerate(header)
                ]

                dict_rows = []
                for row in data_rows:
                    row_dict = {
                        header_safe[i]: row[i] if i < len(row) else ""
                        for i in range(len(header_safe))
                    }
                    dict_rows.append(row_dict)

                schedule_tables.append({
                    "page_index": page_index,
                    "table_index_on_page": t_index,
                    "header": header_safe,
                    "rows": dict_rows,
                })

                # ---- MARK extraction from table cells (FIXED) ----
                for row in data_rows:
                    for cell in row:
                        if cell:
                            m = MARK_REGEX.search(cell)
                            if m:
                                marks_set.add(m.group(0).upper())

            # ---- FALLBACK: scan entire schedule page text ----
            for m in MARK_REGEX.findall(text_upper):
                marks_set.add(m)

    schedule_json = {
        "schedule_tables": schedule_tables,
        "marks": sorted(marks_set),
    }

    return schedule_json, sorted(marks_set)


def mark_type(mark: str) -> str:
    m = re.match(r'^([A-Z]{1,5})', mark)
    return m.group(1).upper() if m else mark.split("-")[0].upper()


def get_plan_label(page: fitz.Page):
    text = page.get_text() or ""
    lines = [l.strip() for l in text.splitlines() if "PLAN" in l.upper()]
    if not lines:
        return None
    for l in lines:
        if "MECHANICAL" in l.upper():
            return l
    return lines[0]


def build_type_color_map(types):
    palette = [
        (1, 0, 0),
        (0, 0, 1),
        (0, 0.6, 0),
        (1, 0.5, 0),
        (0.6, 0, 0.6),
        (0, 0.7, 0.7),
        (0.7, 0.7, 0),
        (0.5, 0.3, 0.1),
        (0, 0, 0.5),
        (0, 0.5, 0.3),
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
                rects.extend(page.search_for(v))

            unique_rects = list({
                (r.x0, r.y0, r.x1, r.y1): r
                for r in rects
            }.values())

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
        type_color_map,
    )


# ---------------- STREAMLIT UI ----------------

st.title("Mechanical Schedules → Mark Highlighter")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    raw_pdf = uploaded_file.read()

    with st.spinner("Making PDF searchable (OCR if needed)..."):
        pdf_bytes = ensure_searchable_pdf(raw_pdf)

    with st.spinner("Extracting schedules and marks..."):
        schedule_json, marks = extract_schedules_and_marks(pdf_bytes)

    st.subheader("Detected Marks")
    st.code(", ".join(marks) if marks else "None", language="text")

    st.subheader("Schedule JSON (preview)")
    st.code(json.dumps(schedule_json, indent=2)[:4000] + "\n...", language="json")

    if marks:
        with st.spinner("Highlighting marks in PDF..."):
            highlighted, plan_mark_counts, plan_type_counts, color_map = highlight_pdf(
                pdf_bytes, marks
            )

        st.download_button(
            "Download highlighted PDF",
            highlighted,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_highlighted.pdf",
            mime="application/pdf",
        )

        st.subheader("Counts per Plan (by mark)")
        st.json(plan_mark_counts)

        st.subheader("Counts per Plan (by type)")
        st.json(plan_type_counts)

        st.subheader("Type → Color map")
        st.json(color_map)
