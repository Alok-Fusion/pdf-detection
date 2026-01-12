import re
import tempfile

import cv2
import fitz
import numpy as np
import streamlit as st
from PIL import Image

st.set_page_config(layout="wide")
st.title("🛠 Auto Mechanical Symbol Detection (Direct from PDF)")

uploaded_pdf = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

# -------------------------------
# Helper: Convert page to image
# -------------------------------
def page_to_image(page, dpi=300):
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)

# -------------------------------
# STEP 1: Extract symbols from legend page
# -------------------------------
def extract_legend_symbols(gray_img):
    _, thresh = cv2.threshold(gray_img, 200, 255, cv2.THRESH_BINARY_INV)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )

    symbols = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)

        # Filter noise (legend symbols have size)
        if 25 < w < 200 and 25 < h < 200:
            symbol = gray_img[y:y+h, x:x+w]
            symbols.append(symbol)

    return symbols

# -------------------------------
# STEP 2: Template matching
# -------------------------------
def match_and_mark(page_img, templates):
    marked = cv2.cvtColor(page_img, cv2.COLOR_GRAY2BGR)

    for tpl in templates:
        h, w = tpl.shape
        if h > page_img.shape[0] or w > page_img.shape[1]:
            continue

        res = cv2.matchTemplate(page_img, tpl, cv2.TM_CCOEFF_NORMED)
        loc = np.where(res > 0.75)

        for pt in zip(*loc[::-1]):
            cv2.rectangle(
                marked,
                pt,
                (pt[0] + w, pt[1] + h),
                (0, 0, 255),
                2
            )

    return marked

# -------------------------------
# MAIN
# -------------------------------
if uploaded_pdf and st.button("🔍 Auto Detect & Mark Symbols"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_pdf.read())
        pdf_path = tmp.name

    doc = fitz.open(pdf_path)

    # -------------------------------
    # Find legend page (M001)
    # -------------------------------
    legend_page = None
    for i, page in enumerate(doc):
        if "MECHANICAL LEGEND" in page.get_text().upper():
            legend_page = page
            break

    if legend_page is None:
        st.error("Mechanical Legend page not found")
        st.stop()

    legend_img = page_to_image(legend_page)

    # -------------------------------
    # Extract symbols automatically
    # -------------------------------
    legend_symbols = extract_legend_symbols(legend_img)
    st.success(f"Extracted {len(legend_symbols)} legend symbols automatically")

    # -------------------------------
    # Process M sheets
    # -------------------------------
    output_images = []

    for page in doc:
        text = page.get_text()
        if not re.search(r"\bM\d{3}\b", text):
            continue

        page_img = page_to_image(page)
        marked_img = match_and_mark(page_img, legend_symbols)
        output_images.append(marked_img)

    # -------------------------------
    # Save output PDF
    # -------------------------------
    out_pdf = pdf_path.replace(".pdf", "_AUTO_SYMBOLS_MARKED.pdf")
    out_doc = fitz.open()

    for img in output_images:
        h, w = img.shape[:2]
        page = out_doc.new_page(width=w, height=h)
        page.insert_image(
            fitz.Rect(0, 0, w, h),
            stream=Image.fromarray(img).tobytes("jpeg", "RGB")
        )

    out_doc.save(out_pdf)
    out_doc.close()
    doc.close()

    with open(out_pdf, "rb") as f:
        st.download_button(
            "⬇ Download Marked PDF",
            f,
            file_name="mechanical_symbols_only.pdf",
            mime="application/pdf"
        )

    st.success("✅ Symbols auto-picked from PDF and marked")
