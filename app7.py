import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import io
import re

# --- CONFIGURATION ---
# UPDATE THIS PATH to your local Poppler bin folder
POPPLER_BIN_PATH = r"C:/poppler-25.12.0/Library/bin"

def smart_scan_and_highlight(pdf_bytes):
    """
    1. Converts PDF to High-Res Images (300 DPI for accuracy).
    2. Uses OCR to read text.
    3. Uses 'Fuzzy Regex' to find tags like AC-1, AC 1, AC- 1.
    """
    try:
        # 300 DPI is critical for reading small blueprint text
        images = convert_from_bytes(pdf_bytes, dpi=300, poppler_path=POPPLER_BIN_PATH)
    except Exception as e:
        st.error(f"Poppler Error: {e}")
        return None, []

    pdf_writer = fitz.open()
    found_tags_log = []
    
    # REGEX PATTERN EXPLAINED:
    # \b          -> Start of word
    # ([A-Z]{1,4})-> 1 to 4 Letters (Type: AC, EF, CU)
    # \s* -> Optional space
    # [-‐‑_]?     -> Optional dash/underscore
    # \s* -> Optional space
    # (\d+)       -> Numbers
    # \b          -> End of word
    # Matches: "AC-1", "AC 1", "AC - 1", "EF10"
    pattern = re.compile(r'\b([A-Z]{1,4})\s*[-‐‑_]?\s*(\d+)\b')

    progress_bar = st.progress(0)
    status = st.empty()

    for i, img in enumerate(images):
        status.text(f"Scanning Page {i+1}/{len(images)}...")
        progress_bar.progress((i + 1) / len(images))

        # 1. OCR (Get text + coordinates)
        try:
            # We get a searchable PDF page directly from Tesseract
            pdf_page_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            page_doc = fitz.open("pdf", pdf_page_bytes)
            page = page_doc[0]
        except Exception as e:
            st.warning(f"OCR failed on page {i+1}: {e}")
            continue

        # 2. Find and Highlight
        # Get all text words to analyze
        text_on_page = page.get_text("text")
        
        # Find all matches in the text
        matches = pattern.finditer(text_on_page)
        
        # Use a set to avoid highlighting the same spot twice
        unique_locs = set()

        for match in matches:
            raw_str = match.group(0) # e.g. "AC - 1"
            tag_type = match.group(1).upper()
            tag_num = match.group(2)
            clean_tag = f"{tag_type}-{tag_num}"

            # Filter: Ignore unlikely tags (e.g., "AT-1" or "IN-2") to reduce noise
            # You can add valid mechanical codes here
            valid_codes = ["AC", "CU", "EF", "HP", "CD", "RG", "SR", "SD", "AH", "VAV", "FCU"]
            if tag_type not in valid_codes:
                continue

            # Search for the *raw string* to get coordinates
            instances = page.search_for(raw_str)

            for inst in instances:
                # Coordinate Key (rounded)
                loc_key = (round(inst.x0), round(inst.y0))
                if loc_key in unique_locs:
                    continue
                unique_locs.add(loc_key)

                # Highlight Color
                # Cyan for Air Distribution, Yellow for Equipment
                is_dist = tag_type in ['CD', 'RG', 'SR', 'SD']
                color = (0, 1, 1) if is_dist else (1, 1, 0)

                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=color)
                annot.update()

                found_tags_log.append({
                    "Page": i + 1,
                    "Tag": clean_tag,
                    "Detected": raw_str
                })

        pdf_writer.insert_pdf(page_doc)

    progress_bar.empty()
    status.empty()
    
    out_buffer = io.BytesIO()
    pdf_writer.save(out_buffer)
    out_buffer.seek(0)
    
    return out_buffer.getvalue(), found_tags_log

# --- UI ---
st.set_page_config(page_title="Precision Tagger")
st.title("🎯 Precision Mechanical Tagger")
st.markdown("Uses **High-Res OCR** + **Smart Pattern Matching** (No AI hallucinations).")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    if "path" in POPPLER_BIN_PATH:
         st.error("⚠️ Please update POPPLER_BIN_PATH in the code!")
         st.stop()
         
    with st.spinner("Processing..."):
        final_pdf, log = smart_scan_and_highlight(uploaded_file.read())
        
    if log:
        st.success(f"Found {len(log)} tags!")
        df = pd.DataFrame(log)
        st.dataframe(df.head(10), use_container_width=True)
        
        st.download_button("📥 Download Highlighted PDF", final_pdf, "marked.pdf", "application/pdf")
    else:
        st.warning("No tags found. Check if your Poppler path is correct.")