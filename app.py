import io
import re
import json
import zipfile
import math

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image
import cv2
import numpy as np
from pdf2image import convert_from_path

# Tesseract path is auto-detected on Streamlit Cloud (installed via packages.txt)

# --------- PAGE CONFIG ---------
st.set_page_config(layout="wide")

# --------- REGEX PATTERNS ---------
MARK_REGEX = re.compile(r'\b[A-Z]{1,4}-\d+\b', re.IGNORECASE)
MEAS_PATTERN = r'\d+"\s*Ø\s*/\s*\d+|\d+"\s*x\s*\d+"\s*/\s*\d+|\d+"\s*x\s*\d+|\d+"\s*Ø'

# ======================================================
# TEXT UTILITIES
# ======================================================

def normalize_text(text: str) -> str:
    return text.replace("—", "-").replace("–", "-")

def extract_marks_from_text(text: str, marks_set: set):
    text = normalize_text(text)
    for m in MARK_REGEX.findall(text):
        marks_set.add(m.upper())

# ======================================================
# OCR LAYER (MAKES PDF FULLY SEARCHABLE)
# ======================================================

def ocr_pdf(pdf_bytes: bytes, target_dpi=300, max_pixels=20000000) -> bytes:
    """
    Memory-safe OCR for large mechanical / architectural PDFs.
    Uses Tesseract to generate searchable PDF pages directly.
    Automatically reduces DPI for very large pages to prevent memory errors.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()

    for page_index, page in enumerate(src):
        try:
            # Calculate safe DPI for this specific page
            rect = page.rect
            width_pts = rect.width
            height_pts = rect.height
            
            # Calculate pixels at target DPI
            width_px = int(width_pts * target_dpi / 72)
            height_px = int(height_pts * target_dpi / 72)
            total_pixels = width_px * height_px
            
            # Adjust DPI if needed
            if total_pixels > max_pixels:
                scale_factor = math.sqrt(max_pixels / total_pixels)
                safe_dpi = max(150, int(target_dpi * scale_factor))
            else:
                safe_dpi = target_dpi
            
            # Render page at safe DPI
            mat = fitz.Matrix(safe_dpi / 72, safe_dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Let Tesseract generate a searchable PDF page
            pdf_bytes_ocr = pytesseract.image_to_pdf_or_hocr(
                img, extension="pdf"
            )

            ocr_page = fitz.open(stream=pdf_bytes_ocr, filetype="pdf")
            out.insert_pdf(ocr_page)

            # Explicit cleanup (important!)
            pix = None
            img = None
            ocr_page.close()
            
        except Exception as e:
            st.warning(f"Could not OCR page {page_index + 1}: {str(e)}. Skipping this page.")
            # Insert blank page to maintain page numbering
            out.insert_pdf(src, from_page=page_index, to_page=page_index)

    buf = io.BytesIO()
    out.save(buf)
    out.close()
    src.close()

    buf.seek(0)
    return buf.getvalue()

# ======================================================
# MARK EXTRACTION & SCHEDULES
# ======================================================

def extract_schedules_and_marks(pdf_bytes: bytes):
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = normalize_text(page.extract_text() or "")
            extract_marks_from_text(text, marks_set)

            for table in page.extract_tables() or []:
                for row in table:
                    for cell in row:
                        if cell:
                            extract_marks_from_text(str(cell), marks_set)

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

            variants = {
                mark,
                mark.replace("-", " "),
                mark.replace("-", ""),
            }

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


def highlight_measurements_on_pdf(pdf_bytes, measurements_df):
    """
    Highlight measurements on PDF with a distinct color (orange).
    Returns the highlighted PDF bytes.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Use orange color for measurements
    measurement_color = (1, 0.5, 0)  # Orange RGB
    
    highlighted_count = 0
    
    for _, row in measurements_df.iterrows():
        page_num = int(row['page']) - 1  # Convert to 0-indexed
        if page_num >= len(doc):
            continue
            
        page = doc[page_num]
        raw_text = row['raw']
        
        # Search for the measurement text
        rects = page.search_for(raw_text)
        
        if rects:
            for rect in rects:
                annot = page.add_highlight_annot(rect)
                annot.set_colors(stroke=measurement_color)
                annot.update()
                highlighted_count += 1
    
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    buf.seek(0)
    
    return buf.getvalue(), highlighted_count


def create_combined_highlighted_pdf(pdf_bytes, marks, measurements_df, file_name):
    """
    Create a single PDF with both marks and measurements highlighted.
    Returns PDF bytes, marks dataframe, and highlight counts.
    """
    # First highlight marks
    pdf_with_marks, marks_df = highlight_pdf_and_collect(pdf_bytes, marks, file_name)
    
    # Then highlight measurements on top
    if not measurements_df.empty:
        combined_pdf, meas_count = highlight_measurements_on_pdf(pdf_with_marks, measurements_df)
    else:
        combined_pdf = pdf_with_marks
        meas_count = 0
    
    return combined_pdf, marks_df, meas_count

# ======================================================
# LEGEND EXTRACTION (VECTOR TEXT)
# ======================================================

def extract_legends(pdf_bytes):
    legends = {}
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page in doc:
        text = page.get_text()
        if "MECHANICAL ABBREVIATIONS" in text.upper():
            for line in text.splitlines():
                m = re.match(r'^([A-Z/]+)\s+(.+)$', line.strip())
                if m:
                    legends[m.group(1)] = m.group(2)
    
    doc.close()
    return legends

# ======================================================
# MECHANICAL PAGE DETECTION
# ======================================================

def find_mechanical_pages(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    mech_pages = []

    for i, page in enumerate(doc):
        text = page.get_text().upper()
        if re.search(r'MP-\d+|MECHANICAL|M&P', text):
            mech_pages.append(i)

    doc.close()
    return mech_pages

# ======================================================
# VECTOR TEXT EXTRACTION (PRIMARY)
# ======================================================

def extract_vector_text(pdf_bytes, pages):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
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

    doc.close()
    return results

# ======================================================
# OCR FALLBACK FOR MEASUREMENTS
# ======================================================

def get_safe_dpi(page, target_dpi=400, max_pixels=20000000):
    """
    Calculate safe DPI that won't exceed memory limits.
    Max 20 million pixels (~4500x4500) to avoid memory errors.
    """
    rect = page.rect
    width_pts = rect.width
    height_pts = rect.height
    
    # Calculate pixels at target DPI
    width_px = int(width_pts * target_dpi / 72)
    height_px = int(height_pts * target_dpi / 72)
    total_pixels = width_px * height_px
    
    # If within limits, use target DPI
    if total_pixels <= max_pixels:
        return target_dpi
    
    # Otherwise, scale down DPI to fit within pixel limit
    scale_factor = math.sqrt(max_pixels / total_pixels)
    safe_dpi = int(target_dpi * scale_factor)
    
    # Minimum 150 DPI for reasonable OCR quality
    return max(150, safe_dpi)


def extract_ocr_text_for_measurements(pdf_bytes, pages):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    ocr_results = []

    for p in pages:
        page = doc[p]
        
        # Calculate safe DPI for this page
        safe_dpi = get_safe_dpi(page, target_dpi=400)
        mat = fitz.Matrix(safe_dpi / 72, safe_dpi / 72)
        
        try:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
            img = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            img = cv2.threshold(img, 170, 255, cv2.THRESH_BINARY)[1]
            
            text = pytesseract.image_to_string(img, config="--psm 6")
            ocr_results.append({
                "page": p + 1,
                "text": text
            })
            
            # Explicit cleanup
            pix = None
            img = None
            
        except Exception as e:
            st.warning(f"Could not OCR page {p + 1} for measurements: {str(e)}")
            ocr_results.append({
                "page": p + 1,
                "text": ""
            })

    doc.close()
    return ocr_results

# ======================================================
# MEASUREMENT PARSER
# ======================================================

def extract_measurements(text):
    return re.findall(MEAS_PATTERN, text)

# ======================================================
# ENGINEERING CALCULATIONS
# ======================================================

def circular_area(d):
    return math.pi * (d / 2) ** 2 / 144

def rectangular_area(w, h):
    return (w * h) / 144

def velocity(cfm, area):
    return round(cfm / area, 1) if cfm and area else None

# ======================================================
# INTERPRET MEASUREMENT
# ======================================================

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

# ======================================================
# STREAMLIT UI
# ======================================================

st.title("🔧 Comprehensive Mechanical PDF Analyzer")
st.markdown("**Extract marks, schedules, measurements, and engineering calculations from mechanical plans**")

uploaded_file = st.file_uploader("Upload Mechanical PDF", type=["pdf"])

# OCR options
st.sidebar.header("⚙️ Processing Options")
run_ocr = st.sidebar.checkbox("Run OCR (for scanned/image PDFs)", value=True, 
                               help="Uncheck if your PDF is already searchable to save time and memory")
if run_ocr:
    ocr_dpi = st.sidebar.slider("OCR Quality (DPI)", min_value=150, max_value=400, value=300, step=50,
                                 help="Lower DPI = faster processing, less memory. Higher DPI = better quality for small text")
else:
    st.sidebar.info("Skipping OCR - will only extract from vector text")

if "processed" not in st.session_state:
    st.session_state.processed = False

if uploaded_file and not st.session_state.processed:
    
    # Read original PDF
    original_pdf = uploaded_file.read()
    
    # Step 1: OCR Processing (optional)
    if run_ocr:
        with st.spinner(f"🔍 Running OCR at {ocr_dpi} DPI to make PDF searchable..."):
            searchable_pdf = ocr_pdf(original_pdf, target_dpi=ocr_dpi)
    else:
        st.info("⏩ Skipping OCR - using original PDF")
        searchable_pdf = original_pdf
    
    # Step 2: Mark Extraction
    with st.spinner("📊 Extracting marks..."):
        marks = extract_schedules_and_marks(searchable_pdf)
        
        if marks:
            marks_json = {
                "file_name": uploaded_file.name,
                "marks_found": len(marks)
            }
        else:
            st.warning("No marks detected in the PDF.")
            marks = []
            marks_json = {}
    
    # Step 3: Mechanical Intelligence Extraction
    with st.spinner("🧠 Extracting mechanical measurements and calculations..."):
        legends = extract_legends(searchable_pdf)
        mech_pages = find_mechanical_pages(searchable_pdf)
        
        if mech_pages:
            vector_text = extract_vector_text(searchable_pdf, mech_pages)
            ocr_text = extract_ocr_text_for_measurements(searchable_pdf, mech_pages)
            
            items = []
            for entry in vector_text + ocr_text:
                for raw in extract_measurements(entry["text"]):
                    items.append(interpret(raw, entry["page"], legends))
            
            measurements_df = pd.DataFrame(items) if items else pd.DataFrame()
        else:
            st.warning("No mechanical pages detected.")
            measurements_df = pd.DataFrame()
    
    # Step 4: Combined Highlighting
    with st.spinner("🎨 Creating combined highlighted PDF (marks + measurements)..."):
        combined_pdf, marks_df, meas_highlight_count = create_combined_highlighted_pdf(
            searchable_pdf, marks, measurements_df, uploaded_file.name
        )
        
        # Update marks JSON with full data
        if not marks_df.empty:
            marks_json = {
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
                    for row in marks_df.itertuples(index=False)
                ],
            }
        else:
            marks_df = pd.DataFrame()
        
        st.success(f"✅ Highlighted {len(marks_df)} mark instances and {meas_highlight_count} measurements!")
    
    # Store in session state
    st.session_state.marks_df = marks_df
    st.session_state.marks_json = marks_json
    st.session_state.measurements_df = measurements_df
    st.session_state.legends = legends
    st.session_state.combined_highlighted_pdf = combined_pdf
    st.session_state.original_pdf = original_pdf
    st.session_state.file_name = uploaded_file.name
    st.session_state.processed = True

# ======================================================
# OUTPUT DISPLAY
# ======================================================

if st.session_state.processed:
    
    # Create tabs for different views
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Marks & Schedules", "📐 Measurements & Calculations", "🏷️ Legends", "⬇️ Download"])
    
    with tab1:
        st.subheader("📊 Extracted Marks Data")
        if not st.session_state.marks_df.empty:
            st.dataframe(st.session_state.marks_df, use_container_width=True, height=450)
            
            st.subheader("🧾 Marks JSON")
            st.json(st.session_state.marks_json)
        else:
            st.info("No marks found in the document.")
    
    with tab2:
        st.subheader("📐 Mechanical Measurements & Engineering Calculations")
        if not st.session_state.measurements_df.empty:
            st.dataframe(st.session_state.measurements_df, use_container_width=True, height=450)
            
            # Engineering summary view
            if "type" in st.session_state.measurements_df.columns:
                st.subheader("📊 Engineering Summary")
                summary_cols = [
                    "page", "symbol", "symbol_meaning", "raw", 
                    "type", "area_ft2", "velocity_fpm"
                ]
                available_cols = [col for col in summary_cols if col in st.session_state.measurements_df.columns]
                st.dataframe(st.session_state.measurements_df[available_cols])
        else:
            st.info("No mechanical measurements detected.")
    
    with tab3:
        st.subheader("🏷️ Mechanical Abbreviations Legend")
        if st.session_state.legends:
            legend_df = pd.DataFrame([
                {"Symbol": k, "Meaning": v} 
                for k, v in st.session_state.legends.items()
            ])
            st.dataframe(legend_df, use_container_width=True)
        else:
            st.info("No legend/abbreviations found in the document.")
    
    with tab4:
        st.subheader("⬇️ Download Results")
        
        # Prepare downloadable files
        csv_marks = st.session_state.marks_df.to_csv(index=False).encode("utf-8") if not st.session_state.marks_df.empty else b""
        csv_measurements = st.session_state.measurements_df.to_csv(index=False).encode("utf-8") if not st.session_state.measurements_df.empty else b""
        json_marks = json.dumps(st.session_state.marks_json, indent=2).encode("utf-8") if st.session_state.marks_json else b""
        
        # Create comprehensive JSON with both marks and measurements
        comprehensive_json = {
            "file_name": st.session_state.file_name,
            "marks": st.session_state.marks_json.get("records", []),
            "measurements": st.session_state.measurements_df.to_dict('records') if not st.session_state.measurements_df.empty else [],
            "legends": st.session_state.legends
        }
        json_comprehensive = json.dumps(comprehensive_json, indent=2).encode("utf-8")
        
        # Create comprehensive ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            # Original and combined highlighted PDF
            zipf.writestr(f"input/{st.session_state.file_name}", st.session_state.original_pdf)
            zipf.writestr(
                f"output/{st.session_state.file_name.rsplit('.',1)[0]}_highlighted_combined.pdf",
                st.session_state.combined_highlighted_pdf
            )
            
            # Data files
            if csv_marks:
                zipf.writestr("data/marks_data.csv", csv_marks)
            if json_marks:
                zipf.writestr("data/marks_data.json", json_marks)
            if csv_measurements:
                zipf.writestr("data/measurements_data.csv", csv_measurements)
            
            # Comprehensive JSON
            zipf.writestr("data/comprehensive_data.json", json_comprehensive)
            
            # Legend
            if st.session_state.legends:
                legend_json = json.dumps(st.session_state.legends, indent=2).encode("utf-8")
                zipf.writestr("data/legends.json", legend_json)
        
        zip_buffer.seek(0)
        
        st.download_button(
            "⬇️ Download Complete Analysis Package (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"{st.session_state.file_name.rsplit('.',1)[0]}_complete_analysis.zip",
            mime="application/zip",
        )
        
        # Individual download buttons
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.download_button(
                "📄 Highlighted PDF",
                data=st.session_state.combined_highlighted_pdf,
                file_name=f"{st.session_state.file_name.rsplit('.',1)[0]}_highlighted.pdf",
                mime="application/pdf"
            )
        
        with col2:
            if csv_marks:
                st.download_button(
                    "📊 Marks CSV",
                    data=csv_marks,
                    file_name="marks_data.csv",
                    mime="text/csv"
                )
        
        with col3:
            if csv_measurements:
                st.download_button(
                    "📐 Measurements CSV",
                    data=csv_measurements,
                    file_name="measurements_data.csv",
                    mime="text/csv"
                )
        
        col4, col5 = st.columns(2)
        
        with col4:
            st.download_button(
                "📋 Comprehensive JSON",
                data=json_comprehensive,
                file_name="comprehensive_data.json",
                mime="application/json"
            )
        
        with col5:
            if st.session_state.legends:
                legend_json = json.dumps(st.session_state.legends, indent=2).encode("utf-8")
                st.download_button(
                    "🏷️ Legends JSON",
                    data=legend_json,
                    file_name="legends.json",
                    mime="application/json"
                )
        
        st.markdown("---")
        st.markdown("**Complete Package Contents:**")
        st.markdown("""
        - `input/` - Original PDF
        - `output/` - **Combined highlighted PDF** (marks in color + measurements in orange)
        - `data/marks_data.csv` - Mark extraction results
        - `data/marks_data.json` - Mark data in JSON format
        - `data/measurements_data.csv` - Mechanical measurements & calculations
        - `data/comprehensive_data.json` - **All data combined** (marks + measurements + legends)
        - `data/legends.json` - Mechanical abbreviations legend
        """)
        
        # Color legend
        st.markdown("---")
        st.markdown("**Highlight Color Legend:**")
        st.markdown("""
        - 🔴 **Red** - Type 1 marks
        - 🔵 **Blue** - Type 2 marks
        - 🟢 **Green** - Type 3 marks
        - 🟠 **Orange** - Measurements (duct sizes, CFM values)
        - 🟣 **Purple** - Additional mark types as needed
        """)
