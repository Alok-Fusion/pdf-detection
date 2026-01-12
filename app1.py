import io
import os
import re
import json
from collections import defaultdict

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd


# ================= CONFIG =================

OUTPUT_ROOT = "output"
MARK_REGEX = re.compile(r"[A-Z]{1,4}-\d+", re.IGNORECASE)
TAG_VALUE_REGEX = re.compile(r"^[A-Z]{1,4}-\d+$", re.IGNORECASE)


# ================= FILE SYSTEM =================

def make_output_dirs(project_name):
    base = os.path.join(OUTPUT_ROOT, project_name)
    highlighted = os.path.join(base, "highlighted")
    data = os.path.join(base, "data")

    os.makedirs(highlighted, exist_ok=True)
    os.makedirs(data, exist_ok=True)

    return {
        "base": base,
        "highlighted": highlighted,
        "data": data,
    }


# ================= EXTRACTION =================

def extract_schedules_and_marks(pdf_bytes):
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for m in MARK_REGEX.findall(text):
                marks_set.add(m.upper())

    return sorted(marks_set)


def read_excel_safely(uploaded_file):
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)

    try:
        if name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        raise ValueError(
            "❌ Unable to read Excel file. Please ensure it is a valid Excel/CSV.\n\n"
            f"Technical error: {e}"
        )

    return df


def auto_detect_tag_column(df):
    # 1️⃣ Header-based
    for col in df.columns:
        if any(k in str(col).upper() for k in ["TAG", "MARK", "EQUIP", "UNIT", "ID"]):
            return col

    # 2️⃣ Value-based
    for col in df.columns:
        values = df[col].dropna().astype(str)
        if len(values) == 0:
            continue

        matches = sum(1 for v in values if TAG_VALUE_REGEX.match(v.strip()))
        if matches / len(values) >= 0.5:
            return col

    return None


# ================= PDF LOGIC =================

def mark_type(mark):
    m = re.match(r"^[A-Z]+", mark)
    return m.group(0) if m else mark


def get_plan_label(page):
    text = page.get_text() or ""
    lines = [l.strip() for l in text.splitlines() if "PLAN" in l.upper()]
    for l in lines:
        if "MECHANICAL" in l.upper():
            return l
    return lines[0] if lines else None


def build_type_color_map(types):
    palette = [
        (1, 0, 0),
        (0, 0, 1),
        (0, 0.6, 0),
        (1, 0.5, 0),
        (0.6, 0, 0.6),
        (0, 0.7, 0.7),
        (0.7, 0.7, 0),
    ]
    return {t: palette[i % len(palette)] for i, t in enumerate(sorted(set(types)))}


def highlight_pdf(pdf_bytes, marks):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    types = [mark_type(m) for m in marks]
    type_color_map = build_type_color_map(types)

    plan_mark_counts = defaultdict(lambda: defaultdict(int))
    plan_type_counts = defaultdict(lambda: defaultdict(int))
    found_tags = set()

    for page in doc:
        plan = get_plan_label(page)

        for mark in marks:
            variants = {mark, mark.replace("-", " "), mark.replace("-", "")}
            rects = []
            for v in variants:
                rects.extend(page.search_for(v))

            rects = list({r for r in rects})
            if not rects:
                continue

            found_tags.add(mark)

            color = type_color_map[mark_type(mark)]
            for r in rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()

            if plan:
                plan_mark_counts[plan][mark] += len(rects)
                plan_type_counts[plan][mark_type(mark)] += len(rects)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    doc.close()
    out_buf.seek(0)

    return (
        out_buf.getvalue(),
        dict(plan_mark_counts),
        dict(plan_type_counts),
        found_tags,
    )


# ================= STREAMLIT UI =================

st.set_page_config(layout="wide")
st.title("📐 Mechanical Schedules → Mark Highlighter")

uploaded_pdf = st.file_uploader("Upload Mechanical PDF", type=["pdf"])
uploaded_excel = st.file_uploader("Upload Excel (optional)", type=["xlsx", "xls", "csv"])

excel_tags = []
excel_df = None

# ---- Excel handling (AUTO) ----
if uploaded_excel:
    try:
        excel_df = read_excel_safely(uploaded_excel)
        tag_col = auto_detect_tag_column(excel_df)

        if tag_col:
            excel_tags = (
                excel_df[tag_col]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
                .unique()
                .tolist()
            )
            st.success(f"TAG column auto-detected: `{tag_col}` ({len(excel_tags)} tags)")
        else:
            st.warning("⚠️ Could not auto-detect TAG column in Excel")

    except Exception as e:
        st.error(str(e))
        st.stop()

run = st.button("🚀 Run Processing")

# ---- Main processing ----
if uploaded_pdf and run:
    pdf_bytes = uploaded_pdf.read()

    if excel_tags:
        all_tags = excel_tags
        st.info("Using tags from Excel")
    else:
        all_tags = extract_schedules_and_marks(pdf_bytes)
        st.info("No Excel provided — using tags from PDF")

    if not all_tags:
        st.error("❌ No tags found to process")
        st.stop()

    highlighted_bytes, plan_mark_counts, plan_type_counts, found_tags = highlight_pdf(
        pdf_bytes, all_tags
    )

    missing_tags = sorted(set(all_tags) - found_tags)

    if missing_tags:
        st.warning("⚠️ Missing Tags (not found on drawings)")
        st.code(", ".join(missing_tags))
    else:
        st.success("✅ All tags found on drawings")

    project = uploaded_pdf.name.rsplit(".", 1)[0]
    dirs = make_output_dirs(project)

    # Save PDF
    with open(os.path.join(dirs["highlighted"], f"{project}_highlighted.pdf"), "wb") as f:
        f.write(highlighted_bytes)

    # Save JSON
    output_json = {
        "tags": all_tags,
        "found_tags": sorted(found_tags),
        "missing_tags": missing_tags,
        "plan_by_mark": plan_mark_counts,
        "plan_by_type": plan_type_counts,
        "excel_used": bool(uploaded_excel),
    }

    with open(os.path.join(dirs["data"], "data.json"), "w") as f:
        json.dump(output_json, f, indent=2)

    # Save Excel
    excel_path = os.path.join(dirs["data"], "summary.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        pd.DataFrame({"TAG": all_tags}).to_excel(writer, "All_Tags", index=False)
        pd.DataFrame({"MISSING_TAG": missing_tags}).to_excel(
            writer, "Missing_Tags", index=False
        )
        pd.DataFrame(plan_mark_counts).T.fillna(0).to_excel(writer, "Plan_by_Tag")
        pd.DataFrame(plan_type_counts).T.fillna(0).to_excel(writer, "Plan_by_Type")
        if excel_df is not None:
            excel_df.to_excel(writer, "Source_Excel", index=False)

    st.success("✅ Outputs saved locally")
    st.code(dirs["base"])

    st.download_button(
        "Download Highlighted PDF",
        highlighted_bytes,
        f"{project}_highlighted.pdf",
        mime="application/pdf",
    )

    st.download_button(
        "Download JSON",
        json.dumps(output_json, indent=2),
        "data.json",
    )

    with open(excel_path, "rb") as f:
        st.download_button("Download Excel Summary", f, "summary.xlsx")
