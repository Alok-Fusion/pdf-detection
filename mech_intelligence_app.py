# # ============================================================
# # Mechanical Plan Intelligence System
# # ============================================================
# # Run with:
# #   streamlit run mech_intelligence_app.py
# #
# # Requirements:
# #   pip install pdfplumber pdf2image pytesseract pillow opencv-python
# #   pip install pandas numpy streamlit
# #
# # NOTE:
# #   Install Tesseract OCR separately and ensure it is in PATH
# # ============================================================

# # ============================================================
# # Mechanical Plan Intelligence System (Detailed Version)
# # ============================================================

# import pdfplumber
# import pytesseract
# import cv2
# import numpy as np
# import re
# import math
# import pandas as pd
# from pdf2image import convert_from_path
# import streamlit as st

# # ============================================================
# # 1. LEGEND EXTRACTION
# # ============================================================

# def extract_legends(pdf_path):
#     legends = {}
#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:
#             text = page.extract_text() or ""
#             if "MECHANICAL ABBREVIATIONS" in text:
#                 for line in text.splitlines():
#                     m = re.match(r'^([A-Z0-9/]+)\s+(.+)$', line.strip())
#                     if m:
#                         legends[m.group(1)] = m.group(2)
#     return legends


# # ============================================================
# # 2. MECHANICAL PAGE DETECTION
# # ============================================================

# def find_mechanical_pages(pdf_path):
#     pages = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for i, page in enumerate(pdf.pages):
#             text = page.extract_text() or ""
#             if re.search(r'M2\.\d|MECHANICAL FLOOR PLAN', text):
#                 pages.append(i)
#     return pages


# # ============================================================
# # 3. OCR
# # ============================================================

# def ocr_pages(pdf_path, pages):
#     images = convert_from_path(pdf_path, dpi=300)
#     results = []

#     for p in pages:
#         img = cv2.cvtColor(np.array(images[p]), cv2.COLOR_BGR2GRAY)
#         img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)[1]
#         text = pytesseract.image_to_string(img, config="--psm 6")
#         results.append((p + 1, text))  # store page number (1-based)

#     return results


# # ============================================================
# # 4. MEASUREMENT EXTRACTION
# # ============================================================

# def extract_measurements(text):
#     pattern = r'\d+"\s*Ø\s*/\s*\d+|\d+"\s*x\s*\d+"\s*/\s*\d+|\d+"\s*x\s*\d+|\d+"\s*Ø'
#     return re.findall(pattern, text)


# # ============================================================
# # 5. ENGINEERING CALCULATIONS
# # ============================================================

# def circular_area(d):
#     return math.pi * (d / 2) ** 2 / 144

# def rectangular_area(w, h):
#     return (w * h) / 144

# def velocity(cfm, area):
#     return round(cfm / area, 1) if cfm and area else None


# # ============================================================
# # 6. INTERPRET MEASUREMENT (WITH SYMBOL + PAGE)
# # ============================================================

# def interpret_measurement(raw, page, legends):
#     data = {
#         "page": page,
#         "raw": raw
#     }

#     # Detect symbol (SA / RA / EA / OA)
#     symbol = None
#     for k in legends:
#         if re.search(rf'\b{k}\b', raw):
#             symbol = k
#             break

#     data["symbol"] = symbol
#     data["symbol_meaning"] = legends.get(symbol)

#     # Circular duct
#     if "Ø" in raw:
#         d = int(re.search(r'(\d+)', raw).group(1))
#         cfm_match = re.search(r'/\s*(\d+)', raw)
#         cfm = int(cfm_match.group(1)) if cfm_match else None

#         area = circular_area(d)
#         data.update({
#             "type": "Circular Duct",
#             "diameter_in": d,
#             "cfm": cfm,
#             "area_ft2": round(area, 3),
#             "velocity_fpm": velocity(cfm, area)
#         })

#     # Rectangular duct
#     if "x" in raw:
#         w, h = map(int, re.findall(r'(\d+)', raw)[:2])
#         cfm_match = re.search(r'/\s*(\d+)', raw)
#         cfm = int(cfm_match.group(1)) if cfm_match else None

#         area = rectangular_area(w, h)
#         data.update({
#             "type": "Rectangular Duct",
#             "width_in": w,
#             "height_in": h,
#             "cfm": cfm,
#             "area_ft2": round(area, 3),
#             "velocity_fpm": velocity(cfm, area)
#         })

#     return data


# # ============================================================
# # 7. STREAMLIT APP
# # ============================================================

# st.set_page_config(layout="wide")
# st.title("🧠 Mechanical Plan Intelligence (Traceable)")

# uploaded = st.file_uploader("Upload Mechanical Plan PDF", type=["pdf"])

# if uploaded:
#     with open("temp.pdf", "wb") as f:
#         f.write(uploaded.read())

#     legends = extract_legends("temp.pdf")
#     pages = find_mechanical_pages("temp.pdf")
#     ocr_results = ocr_pages("temp.pdf", pages)

#     items = []

#     for page_no, text in ocr_results:
#         for raw in extract_measurements(text):
#             items.append(interpret_measurement(raw, page_no, legends))

#     if items:
#         df = pd.DataFrame(items)

#         st.subheader("📊 Detailed BOQ / Takeoff Table")
#         st.dataframe(df)

#         st.subheader("📐 Engineering View")
#         st.dataframe(
#             df[[
#                 "page",
#                 "symbol",
#                 "symbol_meaning",
#                 "raw",
#                 "type",
#                 "area_ft2",
#                 "velocity_fpm"
#             ]]
#         )

#     else:
#         st.warning("No mechanical measurements detected.")


import pdfplumber
import pytesseract
import cv2
import numpy as np
import re
import math
import pandas as pd
from pdf2image import convert_from_path
import streamlit as st

# =====================================================
# LEGEND EXTRACTION (Mechanical only)
# =====================================================

def extract_legends(pdf_path):
    legends = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            if "MECHANICAL SYMBOLS LEGEND" in text or "MECHANICL SYMBOLS LEGEND" in text:
                for line in text.splitlines():
                    line = line.strip()

                    m = re.match(r'(\d+"\s*(?:x|/)?\s*\d*"?\s*Ø?)\s*(.*)', line)

                    if m:
                        size = m.group(1).strip()
                        meaning = m.group(2).strip()

                        if len(meaning) > 2:
                            legends[size] = meaning

    return legends


# =====================================================
# FIND MECHANICAL PLAN PAGES
# =====================================================

def find_mechanical_pages(pdf_path):
    pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""

            if re.search(r'MP-\d+', text):
                pages.append(i)

    return pages


# =====================================================
# OCR + MEASUREMENT EXTRACTION
# =====================================================

def extract_measurements(text):
    pattern = r'\d+"\s*(?:x|/)?\s*\d*"?\s*Ø?(?:\s*/\s*\d+)?'
    return re.findall(pattern, text)


# =====================================================
# ENGINEERING CALCULATIONS
# =====================================================

def circular_area(d):
    return math.pi * (d / 2) ** 2 / 144

def rectangular_area(w, h):
    return (w * h) / 144

def velocity(cfm, area):
    return round(cfm / area, 1) if cfm and area else None


# =====================================================
# INTERPRET MEASUREMENT
# =====================================================

def interpret_measurement(raw, page, legends):
    data = {"page": page, "raw": raw}

    system = "UNKNOWN"

    for key, meaning in legends.items():
        if key in raw:
            data["symbol"] = key
            data["symbol_meaning"] = meaning

            if "S/A" in meaning: system = "SUPPLY AIR"
            elif "R/A" in meaning: system = "RETURN AIR"
            elif "E/A" in meaning: system = "EXHAUST AIR"
            elif "O/A" in meaning: system = "OUTSIDE AIR"
            elif "GE/A" in meaning: system = "GREASE EXHAUST"
            elif "SE/A" in meaning: system = "SMOKE EXHAUST"
            elif "L/A" in meaning: system = "RELIEF AIR"

            break

    data["system"] = system

    if "Ø" in raw:
        d = int(re.search(r'(\d+)', raw).group(1))
        cfm_match = re.search(r'/\s*(\d+)', raw)
        cfm = int(cfm_match.group(1)) if cfm_match else None

        area = circular_area(d)

        data.update({
            "type": "Round",
            "diameter_in": d,
            "cfm": cfm,
            "area_ft2": round(area, 3),
            "velocity_fpm": velocity(cfm, area)
        })

    elif "x" in raw or "/" in raw:
        nums = list(map(int, re.findall(r'\d+', raw)))

        if len(nums) >= 2:
            w, h = nums[:2]
            cfm_match = re.search(r'/\s*(\d+)', raw)
            cfm = int(cfm_match.group(1)) if cfm_match else None

            area = rectangular_area(w, h)

            data.update({
                "type": "Rectangular",
                "width_in": w,
                "height_in": h,
                "cfm": cfm,
                "area_ft2": round(area, 3),
                "velocity_fpm": velocity(cfm, area)
            })

    return data


# =====================================================
# VISUAL HIGHLIGHTING
# =====================================================

def highlight_measurements(image, measurements):
    img = image.copy()

    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    for i, word in enumerate(data["text"]):
        for m in measurements:
            if word in m:
                x = data["left"][i]
                y = data["top"][i]
                w = data["width"][i]
                h = data["height"][i]

                cv2.rectangle(img, (x, y), (x+w, y+h), (0, 0, 255), 2)

    return img


# =====================================================
# STREAMLIT APP
# =====================================================

st.set_page_config(layout="wide")
st.title("🧠 Mechanical Plan Intelligence + System BOQ")

uploaded = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

if uploaded:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded.read())

    legends = extract_legends("temp.pdf")
    pages = find_mechanical_pages("temp.pdf")
    images = convert_from_path("temp.pdf", dpi=300)

    items = []
    highlighted = []

    for p in pages:
        img = cv2.cvtColor(np.array(images[p]), cv2.COLOR_BGR2GRAY)
        img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)[1]

        text = pytesseract.image_to_string(img, config="--psm 6")

        found = extract_measurements(text)

        if found:
            marked = highlight_measurements(img, found)
            highlighted.append((p+1, marked))

        for raw in found:
            items.append(interpret_measurement(raw, p+1, legends))

    if items:
        df = pd.DataFrame(items)

        st.subheader("📋 Full Extracted Takeoff")
        st.dataframe(df)

        # =============================
        # SYSTEM-WISE BOQ
        # =============================

        boq = (
            df.groupby(["system", "type"])
              .agg(
                  count=("raw", "count"),
                  total_area_ft2=("area_ft2", "sum"),
                  total_cfm=("cfm", "sum")
              )
              .reset_index()
        )

        st.subheader("📊 System-wise BOQ Summary")
        st.dataframe(boq)

        # =============================
        # HIGHLIGHTED DRAWINGS
        # =============================

        st.subheader("📌 Highlighted Mechanical Plans")

        for page_no, img in highlighted:
            st.image(img, caption=f"Mechanical Page {page_no}", use_column_width=True)

    else:
        st.warning("No duct measurements detected.")
