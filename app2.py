import io
import os
import cv2
import numpy as np
import streamlit as st
import fitz
from pdf2image import convert_from_bytes
from PIL import Image


# ================= CONFIG =================

OUTPUT_ROOT = "output"
MATCH_THRESHOLD = 0.75  # lower = more tolerant


# ================= FILE SYSTEM =================

def make_output_dirs(project):
    base = os.path.join(OUTPUT_ROOT, project)
    os.makedirs(base, exist_ok=True)
    return base


# ================= LEGEND DETECTION =================

def find_legend_page(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for i, page in enumerate(doc):
        text = (page.get_text() or "").upper()
        if "LEGEND" in text and "SYMBOL" in text:
            doc.close()
            return i
    doc.close()
    return None


# ================= SYMBOL CROPPING =================

def extract_legend_symbol_images(legend_img: Image.Image):
    """
    Very robust for tabular legends like yours.
    """
    w, h = legend_img.size

    # Left column where symbols are (based on your PDF)
    symbol_col = legend_img.crop((0, int(0.15*h), int(0.25*w), h))

    symbols = []
    row_h = 80

    for y in range(0, symbol_col.height - row_h, row_h):
        crop = symbol_col.crop((0, y, symbol_col.width, y + row_h))
        gray = cv2.cvtColor(np.array(crop), cv2.COLOR_BGR2GRAY)

        if gray.mean() < 245:  # ignore empty rows
            symbols.append(gray)

    return symbols


# ================= TEMPLATE MATCH =================

def symbol_exists_on_page(template, page_gray):
    res = cv2.matchTemplate(page_gray, template, cv2.TM_CCOEFF_NORMED)
    return np.max(res) >= MATCH_THRESHOLD


# ================= MAIN PROCESS =================

def tick_legend_symbols(pdf_bytes):
    images = convert_from_bytes(pdf_bytes, dpi=300)

    legend_idx = find_legend_page(pdf_bytes)
    if legend_idx is None:
        raise ValueError("Legend page not found")

    legend_symbols = extract_legend_symbol_images(images[legend_idx])

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = [False] * len(legend_symbols)

    for i, page in enumerate(doc):
        if i == legend_idx:
            continue

        page_img = images[i]
        page_gray = cv2.cvtColor(np.array(page_img), cv2.COLOR_BGR2GRAY)

        for idx, tpl in enumerate(legend_symbols):
            if not found[idx] and symbol_exists_on_page(tpl, page_gray):
                found[idx] = True

    # Tick legend
    legend_page = doc[legend_idx]
    y = 150
    for idx, is_found in enumerate(found):
        if is_found:
            legend_page.insert_text(
                fitz.Point(300, y + idx * 80),
                "✔",
                fontsize=16,
                color=(0, 0.6, 0),
            )

    out = io.BytesIO()
    doc.save(out)
    doc.close()
    out.seek(0)

    return out.getvalue(), sum(found), len(found)


# ================= STREAMLIT =================

st.set_page_config(layout="wide")
st.title("🛠 Mechanical Legend Symbol Checker (WORKING)")

uploaded_pdf = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

if uploaded_pdf and st.button("Run"):
    try:
        pdf_bytes = uploaded_pdf.read()

        with st.spinner("Detecting symbols from legend..."):
            final_pdf, found, total = tick_legend_symbols(pdf_bytes)

        project = uploaded_pdf.name.rsplit(".", 1)[0]
        out_dir = make_output_dirs(project)

        with open(os.path.join(out_dir, "legend_checked.pdf"), "wb") as f:
            f.write(final_pdf)

        st.success(f"✅ {found} / {total} symbols detected in plans")

        st.download_button(
            "Download Result PDF",
            final_pdf,
            "legend_checked.pdf",
            mime="application/pdf",
        )

    except Exception as e:
        st.error(str(e))
