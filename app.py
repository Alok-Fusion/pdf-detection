import io
import re
import json
import zipfile

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd


# --------- PAGE CONFIG ---------
st.set_page_config(layout="wide")


# --------- REGEX ---------
MARK_REGEX = re.compile(r'^[A-Z]{1,4}-\d+', re.IGNORECASE)


# --------- CORE LOGIC ---------

def extract_schedules_and_marks(pdf_bytes: bytes):
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = (page.extract_text() or "")
            if "SCHEDULE" not in text.upper():
                continue

            for table in page.extract_tables() or []:
                for row in table:
                    for cell in row:
                        if cell:
                            m = MARK_REGEX.match(str(cell).strip())
                            if m:
                                marks_set.add(m.group(0).upper())

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

            variants = {mark, mark.replace("-", " "), mark.replace("-", "")}
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


# --------- STREAMLIT UI ---------

st.title("Mechanical PDF → CSV + JSON Viewer")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if "processed" not in st.session_state:
    st.session_state.processed = False


if uploaded_file and not st.session_state.processed:
    with st.spinner("Processing PDF (one-time)..."):
        pdf_bytes = uploaded_file.read()
        marks = extract_schedules_and_marks(pdf_bytes)

        if not marks:
            st.error("No marks detected.")
            st.stop()

        highlighted_pdf, master_df = highlight_pdf_and_collect(
            pdf_bytes, marks, uploaded_file.name
        )

        # ---- JSON STRUCTURE ----
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
        st.session_state.original_pdf = pdf_bytes
        st.session_state.file_name = uploaded_file.name
        st.session_state.processed = True


# --------- DISPLAY ---------

if st.session_state.processed:
    st.subheader("📊 Extracted Data (Table)")
    st.dataframe(
        st.session_state.master_df,
        use_container_width=True,
        height=450,
    )

    st.subheader("🧾 Extracted Data (JSON)")
    st.json(st.session_state.master_json)

    # --------- EXPORT ---------

    csv_bytes = st.session_state.master_df.to_csv(index=False).encode("utf-8")
    json_bytes = json.dumps(
        st.session_state.master_json, indent=2
    ).encode("utf-8")

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
        "⬇️ Download ZIP (CSV + JSON + PDFs)",
        data=zip_buffer.getvalue(),
        file_name=f"{st.session_state.file_name.rsplit('.',1)[0]}_results.zip",
        mime="application/zip",
    )
