import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_bytes
import io
import re
import json
from collections import defaultdict

# --- CONFIGURATION ---
# UPDATE THIS PATH to your Poppler bin folder
POPPLER_BIN_PATH = r"C:/Alok/poppler-25.12.0/Library/bin"

# --- 1. OCR ENGINE ---
def create_searchable_pdf(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=200, poppler_path=POPPLER_BIN_PATH)
    except Exception as e:
        st.error(f"Poppler Error: {e}")
        st.stop()

    pdf_writer = fitz.open()
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, img in enumerate(images):
        status_text.text(f"OCR: Reading Page {i+1}/{len(images)}...")
        progress_bar.progress((i + 1) / len(images))
        try:
            pdf_page_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            img_pdf = fitz.open("pdf", pdf_page_bytes)
            pdf_writer.insert_pdf(img_pdf)
        except Exception as e:
            st.warning(f"OCR Warning on page {i+1}: {e}")

    progress_bar.empty()
    status_text.empty()
    
    out_buffer = io.BytesIO()
    pdf_writer.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue()

# --- 2. ROBUST EXTRACTION (Table + Text Fallback) ---
# Regex: Matches "AC-1", "AC 1", "AC - 1"
# It finds the pattern anywhere, not just in tables.
MARK_REGEX_BROAD = re.compile(r'\b([A-Z]{1,4})\s*[-_]\s*(\d+)\b', re.IGNORECASE)

def extract_marks_robust(pdf_bytes: bytes):
    schedule_data = []
    marks_set = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = (page.extract_text() or "")
            
            # 1. TRY TABLE EXTRACTION FIRST (Best for specs)
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Quick scan of table cells for marks
                    for row in table:
                        for cell in row:
                            if cell:
                                # Clean regex match
                                m = MARK_REGEX_BROAD.search(str(cell))
                                if m:
                                    tag = f"{m.group(1).upper()}-{m.group(2)}"
                                    marks_set.add(tag)
                                    
            # 2. FALLBACK: RAW TEXT SCAN (Best for OCR)
            # If we didn't find much, or just to be safe, scan the whole text block
            text_matches = MARK_REGEX_BROAD.findall(text)
            for (tag_type, tag_num) in text_matches:
                full_tag = f"{tag_type.upper()}-{tag_num}"
                marks_set.add(full_tag)
                
            # Keep track of text for debugging
            if "SCHEDULE" in text.upper() or len(text_matches) > 0:
                schedule_data.append({
                    "page": page_index + 1,
                    "tags_found": len(text_matches),
                    "snippet": text[:200] + "..." # Preview text
                })

    return {"debug_info": schedule_data, "marks": sorted(marks_set)}, sorted(marks_set)

def mark_type(mark: str) -> str:
    return mark.split('-')[0]

def build_type_color_map(types: list[str]):
    palette = [
        (1, 0, 0), (0, 0, 1), (0, 0.6, 0), (1, 0.5, 0), 
        (0.6, 0, 0.6), (0, 0.7, 0.7), (0.7, 0.7, 0), 
        (0.5, 0.3, 0.1), (0, 0, 0.5)
    ]
    return {t: palette[idx % len(palette)] for idx, t in enumerate(sorted(set(types)))}

# --- 3. HIGHLIGHTING ---
def highlight_pdf(pdf_bytes: bytes, marks: list[str]):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    types = [mark_type(m) for m in marks]
    type_color_map = build_type_color_map(types)
    
    found_counts = defaultdict(int)

    for page in doc:
        for mark in marks:
            t = mark_type(mark)
            color = type_color_map.get(t, (1, 1, 0))
            
            # Robust Search variants
            variants = {mark, mark.replace("-", " "), mark.replace("-", "")}
            
            rects = []
            for v in variants:
                rects += page.search_for(v, quads=False)
            
            # Deduplicate
            unique_rects = []
            for r in rects:
                if not any(r == u for u in unique_rects):
                    unique_rects.append(r)
            
            for r in unique_rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()
                found_counts[mark] += 1

    out_buf = io.BytesIO()
    doc.save(out_buf)
    doc.close()
    out_buf.seek(0)
    
    return out_buf.getvalue(), found_counts

# --- APP UI ---
st.set_page_config(page_title="Robust Tag Highlighter", layout="wide")
st.title("🚀 Robust OCR & Highlighter")
st.markdown("""
**New Strategy:**
1.  **OCR:** Make PDF readable.
2.  **Dual Extraction:** Scans **Tables** AND **Raw Text**. (Fixes "No marks found").
3.  **Highlight:** Maps them back to the plan.
""")

if "path" in POPPLER_BIN_PATH:
    st.error("⚠️ **SETUP:** Update `POPPLER_BIN_PATH` in the code!")
    st.stop()

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    raw_bytes = uploaded_file.read()

    # 1. OCR
    st.info("Step 1: Running OCR... (Reading the scan)")
    searchable_bytes = create_searchable_pdf(raw_bytes)
    st.success("OCR Complete.")

    # 2. Extract
    with st.spinner("Step 2: Scanning for tags (Table + Text Search)..."):
        debug_json, marks = extract_marks_robust(searchable_bytes)

    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader(f"Found {len(marks)} Unique Tags")
        if marks:
            st.write(marks)
        else:
            st.error("Still no marks found.")
            st.write("Debug Info (What text did OCR see?):")
            st.json(debug_json) # Shows us if OCR text is garbage

    with col2:
        if marks:
            st.subheader("Highlighter")
            with st.spinner("Step 3: Highlighting..."):
                final_pdf, counts = highlight_pdf(searchable_bytes, marks)
            
            st.download_button("📥 Download Result", final_pdf, "highlighted_robust.pdf", "application/pdf")
            st.write("Highlight Counts:", counts)