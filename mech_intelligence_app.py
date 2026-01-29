# ============================================================
# Mechanical Plan Intelligence System
# ============================================================
# Run with:
#   streamlit run mech_intelligence_app.py
#
# Requirements:
#   pip install pdfplumber pdf2image pytesseract pillow opencv-python
#   pip install pandas numpy streamlit
#
# NOTE:
#   Install Tesseract OCR separately and ensure it is in PATH
# ============================================================

# ============================================================
# Mechanical Plan Intelligence System (Detailed Version)
# ============================================================

import pdfplumber
import pytesseract
import cv2
import numpy as np
import re
import math
import pandas as pd
from pdf2image import convert_from_path
import streamlit as st

# ============================================================
# 1. LEGEND EXTRACTION
# ============================================================

def extract_legends(pdf_path):
    legends = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "MECHANICAL ABBREVIATIONS" in text:
                for line in text.splitlines():
                    m = re.match(r'^([A-Z0-9/]+)\s+(.+)$', line.strip())
                    if m:
                        legends[m.group(1)] = m.group(2)
    return legends


# ============================================================
# 2. MECHANICAL PAGE DETECTION
# ============================================================

def find_mechanical_pages(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if re.search(r'M2\.\d|MECHANICAL FLOOR PLAN', text):
                pages.append(i)
    return pages


# ============================================================
# 3. OCR
# ============================================================

def ocr_pages(pdf_path, pages):
    images = convert_from_path(pdf_path, dpi=300)
    results = []

    for p in pages:
        img = cv2.cvtColor(np.array(images[p]), cv2.COLOR_BGR2GRAY)
        img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)[1]
        text = pytesseract.image_to_string(img, config="--psm 6")
        results.append((p + 1, text))  # store page number (1-based)

    return results


# ============================================================
# 4. MEASUREMENT EXTRACTION
# ============================================================

def extract_measurements(text):
    pattern = r'\d+"\s*Ø\s*/\s*\d+|\d+"\s*x\s*\d+"\s*/\s*\d+|\d+"\s*x\s*\d+|\d+"\s*Ø'
    return re.findall(pattern, text)


# ============================================================
# 5. ENGINEERING CALCULATIONS
# ============================================================

def circular_area(d):
    return math.pi * (d / 2) ** 2 / 144

def rectangular_area(w, h):
    return (w * h) / 144

def velocity(cfm, area):
    return round(cfm / area, 1) if cfm and area else None

# ============================================================
# 6. INTERPRET MEASUREMENT (WITH SYMBOL + PAGE)
# ============================================================

def interpret_measurement(raw, page, legends):
    data = {
        "page": page,
        "raw": raw
    }

    # Detect symbol (SA / RA / EA / OA)
    symbol = None
    for k in legends:
        if re.search(rf'\b{k}\b', raw):
            symbol = k
            break

    data["symbol"] = symbol
    data["symbol_meaning"] = legends.get(symbol)

    # Circular duct
    if "Ø" in raw:
        d = int(re.search(r'(\d+)', raw).group(1))
        cfm_match = re.search(r'/\s*(\d+)', raw)
        cfm = int(cfm_match.group(1)) if cfm_match else None

        area = circular_area(d)
        data.update({
            "type": "Circular Duct",
            "diameter_in": d,
            "cfm": cfm,
            "area_ft2": round(area, 3),
            "velocity_fpm": velocity(cfm, area)
        })

    # Rectangular duct
    if "x" in raw:
        w, h = map(int, re.findall(r'(\d+)', raw)[:2])
        cfm_match = re.search(r'/\s*(\d+)', raw)
        cfm = int(cfm_match.group(1)) if cfm_match else None

        area = rectangular_area(w, h)
        data.update({
            "type": "Rectangular Duct",
            "width_in": w,
            "height_in": h,
            "cfm": cfm,
            "area_ft2": round(area, 3),
            "velocity_fpm": velocity(cfm, area)
        })

    return data


# ============================================================
# 7. STREAMLIT APP
# ============================================================

st.set_page_config(layout="wide")
st.title("🧠 Mechanical Plan Intelligence (Traceable)")

uploaded = st.file_uploader("Upload Mechanical Plan PDF", type=["pdf"])

if uploaded:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded.read())

    legends = extract_legends("temp.pdf")
    pages = find_mechanical_pages("temp.pdf")
    ocr_results = ocr_pages("temp.pdf", pages)

    items = []

    for page_no, text in ocr_results:
        for raw in extract_measurements(text):
            items.append(interpret_measurement(raw, page_no, legends))

    if items:
        df = pd.DataFrame(items)

        st.subheader("📊 Detailed BOQ / Takeoff Table")
        st.dataframe(df)

        st.subheader("📐 Engineering View")
        st.dataframe(
            df[[
                "page",
                "symbol",
                "symbol_meaning",
                "raw",
                "type",
                "area_ft2",
                "velocity_fpm"
            ]]
        )

    else:
        st.warning("No mechanical measurements detected.")
