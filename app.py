import io
import re
import json
from collections import defaultdict

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF


# --------- UTILITIES ---------

MARK_REGEX = re.compile(r'^[A-Z]{1,4}-\d+', re.IGNORECASE)


def extract_schedules_and_marks(pdf_bytes: bytes):
    """
    Extract schedule-like tables and marks from a PDF.
    Returns:
      schedule_json: dict with tables
      marks: sorted list of unique mark strings (e.g., 'FCU-1')
    """
    schedule_tables = []
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = (page.extract_text() or "")
            text_upper = text.upper()

            # Heuristic: pages with "SCHEDULE"
            if "SCHEDULE" not in text_upper:
                continue

            tables = page.extract_tables()
            for t_index, table in enumerate(tables or []):
                if not table:
                    continue

                # Clean table rows
                cleaned = [
                    [("" if c is None else str(c).strip()) for c in row]
                    for row in table
                ]

                # Skip completely empty tables
                if all(all(cell == "" for cell in row) for row in cleaned):
                    continue

                # Guess header row: first row with at least one non-empty cell
                header_row_idx = None
                for i, row in enumerate(cleaned):
                    if any(cell != "" for cell in row):
                        header_row_idx = i
                        break

                if header_row_idx is None:
                    continue

                header = cleaned[header_row_idx]
                data_rows = [
                    row for row in cleaned[header_row_idx + 1 :]
                    if any(cell != "" for cell in row)
                ]

                if not data_rows:
                    continue

                # Build table dict
                header_safe = [
                    col if col else f"COL_{idx+1}" for idx, col in enumerate(header)
                ]

                dict_rows = []
                for row in data_rows:
                    row_dict = {}
                    for col_idx, col_name in enumerate(header_safe):
                        value = row[col_idx] if col_idx < len(row) else ""
                        row_dict[col_name] = value
                    dict_rows.append(row_dict)

                schedule_tables.append(
                    {
                        "page_index": page_index,
                        "table_index_on_page": t_index,
                        "header": header_safe,
                        "rows": dict_rows,
                    }
                )

                # --- Extract marks via regex on each cell ---
                for row in data_rows:
                    for cell in row:
                        if not cell:
                            continue
                        m = MARK_REGEX.match(cell)
                        if m:
                            # Normalize to upper-case + standard hyphen
                            marks_set.add(m.group(0).upper())

    schedule_json = {
        "schedule_tables": schedule_tables,
        "marks": sorted(marks_set),
    }
    return schedule_json, sorted(marks_set)


def mark_type(mark: str) -> str:
    """
    Get type prefix from a mark, e.g. 'FCU-1' -> 'FCU'.
    """
    m = re.match(r'^([A-Z]{1,4})', mark, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    if '-' in mark:
        return mark.split('-')[0].upper()
    return mark.upper()


def get_plan_label(page: fitz.Page) -> str | None:
    """
    Find a 'plan label' on a page, e.g. 'MECHANICAL FLOOR PLAN'.
    """
    text = page.get_text() or ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    plan_lines = [l for l in lines if "PLAN" in l.upper()]
    if not plan_lines:
        return None

    # Prefer lines with 'MECHANICAL' or 'FLOOR/ROOF'
    preferred = [l for l in plan_lines if "MECHANICAL" in l.upper()]
    if preferred:
        return preferred[0]
    for key in ("FLOOR PLAN", "ROOF PLAN", "PIPING PLAN", "DETECTION PLAN"):
        for l in plan_lines:
            if key in l.upper():
                return l
    return plan_lines[0]


def build_type_color_map(types: list[str]):
    """
    Assign a distinct RGB color (0-1) per mark type.
    """
    palette = [
        (1, 0, 0),        # red
        (0, 0, 1),        # blue
        (0, 0.6, 0),      # green
        (1, 0.5, 0),      # orange
        (0.6, 0, 0.6),    # purple
        (0, 0.7, 0.7),    # teal
        (0.7, 0.7, 0),    # yellow-ish
        (0.5, 0.3, 0.1),  # brown
        (0, 0, 0.5),      # dark blue
        (0, 0.5, 0.3),    # dark green
        (0.5, 0.5, 0.5),  # gray
    ]
    type_color_map = {}
    for idx, t in enumerate(sorted(set(types))):
        type_color_map[t] = palette[idx % len(palette)]
    return type_color_map


def highlight_pdf(pdf_bytes: bytes, marks: list[str]):
    """
    Highlight all marks in the PDF, color-coded by type.
    Returns:
      highlighted_pdf_bytes,
      plan_mark_counts,
      plan_type_counts,
      type_color_map
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    types = [mark_type(m) for m in marks]
    type_color_map = build_type_color_map(types)

    plan_mark_counts = defaultdict(lambda: defaultdict(int))
    plan_type_counts = defaultdict(lambda: defaultdict(int))

    for page in doc:
        plan_label = get_plan_label(page)
        # We'll still highlight on non-plan pages, but only count on plan pages

        for mark in marks:
            if not mark:
                continue

            t = mark_type(mark)
            color = type_color_map.get(t, (1, 1, 0))

            # Try a couple of search variants
            variants = {
                mark,
                mark.replace("-", " "),   # FCU 1
                mark.replace("-", ""),    # FCU1
            }

            rects = []
            for v in variants:
                rects += page.search_for(v, quads=False)

            if not rects:
                # Here is where you could add OCR fallback:
                #   - render page to image
                #   - run pytesseract
                #   - find approximate positions
                # For now we just skip.
                continue

            # Deduplicate overlapping rects
            unique_rects = []
            for r in rects:
                if not any(r == u for u in unique_rects):
                    unique_rects.append(r)

            for r in unique_rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()

            if plan_label:
                plan_mark_counts[plan_label][mark] += len(unique_rects)
                plan_type_counts[plan_label][t] += len(unique_rects)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    doc.close()
    out_buf.seek(0)

    # Convert counts to plain dicts
    plan_mark_counts = {p: dict(m) for p, m in plan_mark_counts.items()}
    plan_type_counts = {p: dict(m) for p, m in plan_type_counts.items()}

    return out_buf.getvalue(), plan_mark_counts, plan_type_counts, type_color_map


# --------- STREAMLIT UI ---------

st.title("Mechanical Schedules → Mark Highlighter")

st.write(
    "Upload a mechanical PDF set (with schedule + plan sheets). "
    "The app will extract **marks** from schedules, highlight them on all pages, "
    "and show counts by plan and by mark type."
)

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()

    with st.spinner("Extracting schedules and marks..."):
        schedule_json, marks = extract_schedules_and_marks(pdf_bytes)

    st.subheader("Detected Marks")
    if marks:
        st.code(", ".join(marks), language="text")
    else:
        st.warning("No marks detected in schedule pages.")

    st.subheader("Schedule JSON (preview)")
    st.code(json.dumps(schedule_json, indent=2)[:4000] + "\n...", language="json")

    if marks:
        with st.spinner("Highlighting marks in PDF..."):
            highlighted_bytes, plan_mark_counts, plan_type_counts, type_color_map = highlight_pdf(
                pdf_bytes, marks
            )

        st.subheader("Download Highlighted PDF")
        st.download_button(
            label="Download highlighted PDF",
            data=highlighted_bytes,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_highlighted.pdf",
            mime="application/pdf",
        )

        st.subheader("Counts per Plan (by mark)")
        st.json(plan_mark_counts)

        st.subheader("Counts per Plan (by type)")
        st.json(plan_type_counts)

        st.subheader("Type → Color map")
        st.json(type_color_map)
