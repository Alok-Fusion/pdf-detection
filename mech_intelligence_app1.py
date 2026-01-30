# ============================================================
# Mechanical Plan Intelligence System (VECTOR-FIRST FIX)
# Works on CAD / Revit exported PDFs
# ============================================================

import fitz  # PyMuPDF
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
# 1. LEGEND EXTRACTION (VECTOR TEXT)
# ============================================================

def extract_legends(pdf_path):
    legends = {}
    doc = fitz.open(pdf_path)

    for page in doc:
        text = page.get_text()
        if "MECHANICAL ABBREVIATIONS" in text.upper():
            for line in text.splitlines():
                m = re.match(r'^([A-Z/]+)\s+(.+)$', line.strip())
                if m:
                    legends[m.group(1)] = m.group(2)
    return legends


# ============================================================
# 2. MECHANICAL PAGE DETECTION (ROBUST)
# ============================================================

def find_mechanical_pages(pdf_path):
    doc = fitz.open(pdf_path)
    mech_pages = []

    for i, page in enumerate(doc):
        text = page.get_text().upper()
        if re.search(r'MP-\d+|MECHANICAL|M&P', text):
            mech_pages.append(i)

    return mech_pages


# ============================================================
# 3. VECTOR TEXT EXTRACTION (PRIMARY)
# ============================================================

def extract_vector_text(pdf_path, pages):
    doc = fitz.open(pdf_path)
    results = []

    for p in pages:
        page = doc[p]
        blocks = page.get_text("blocks")

        for b in blocks:
            txt = b[4].strip()
            if txt:
                results.append({
                    "page": p + 1,
                    "text": txt
                })

    return results


# ============================================================
# 4. OCR FALLBACK (ONLY IF NEEDED)
# ============================================================

def extract_ocr_text(pdf_path, pages):
    images = convert_from_path(pdf_path, dpi=400)
    ocr_results = []

    for p in pages:
        img = cv2.cvtColor(np.array(images[p]), cv2.COLOR_BGR2GRAY)
        img = cv2.threshold(img, 170, 255, cv2.THRESH_BINARY)[1]
        text = pytesseract.image_to_string(img, config="--psm 6")
        ocr_results.append({
            "page": p + 1,
            "text": text
        })

    return ocr_results


# ============================================================
# 5. MEASUREMENT PARSER (CAD-AWARE)
# ============================================================

MEAS_PATTERN = r'\d+"\s*Ø\s*/\s*\d+|\d+"\s*x\s*\d+"\s*/\s*\d+|\d+"\s*x\s*\d+|\d+"\s*Ø'

def extract_measurements(text):
    return re.findall(MEAS_PATTERN, text)


# ============================================================
# 6. ENGINEERING CALCULATIONS
# ============================================================

def circular_area(d):
    return math.pi * (d / 2) ** 2 / 144

def rectangular_area(w, h):
    return (w * h) / 144

def velocity(cfm, area):
    return round(cfm / area, 1) if cfm and area else None


# ============================================================
# 7. INTERPRET MEASUREMENT (TRACEABLE)
# ============================================================

def interpret(raw, page, legends):
    data = {
        "page": page,
        "raw": raw
    }

    # System symbol detection
    symbol = None
    for k in legends:
        if k in raw:
            symbol = k
            break

    data["symbol"] = symbol
    data["symbol_meaning"] = legends.get(symbol)

    if "Ø" in raw:
        d = int(re.search(r'(\d+)', raw).group(1))
        cfm = re.search(r'/\s*(\d+)', raw)
        cfm = int(cfm.group(1)) if cfm else None
        area = circular_area(d)

        data.update({
            "type": "Circular Duct",
            "diameter_in": d,
            "cfm": cfm,
            "area_ft2": round(area, 3),
            "velocity_fpm": velocity(cfm, area)
        })

    if "x" in raw:
        w, h = map(int, re.findall(r'(\d+)', raw)[:2])
        cfm = re.search(r'/\s*(\d+)', raw)
        cfm = int(cfm.group(1)) if cfm else None
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
# 8. STREAMLIT APP
# ============================================================

st.set_page_config(layout="wide")
st.title("🧠 Mechanical Plan Intelligence (Vector-Fixed)")

uploaded = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

if uploaded:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded.read())

    legends = extract_legends("temp.pdf")
    mech_pages = find_mechanical_pages("temp.pdf")

    vector_text = extract_vector_text("temp.pdf", mech_pages)
    ocr_text = extract_ocr_text("temp.pdf", mech_pages)

    items = []

    for entry in vector_text + ocr_text:
        for raw in extract_measurements(entry["text"]):
            items.append(interpret(raw, entry["page"], legends))

    if items:
        df = pd.DataFrame(items)

        st.subheader("📊 Detailed Takeoff Table")
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
        st.error("No mechanical measurements detected. Drawing may be symbol-only.")
