import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import io
import re

# --- CONFIGURATION ---
# UPDATE THIS PATH to your Poppler bin folder
POPPLER_BIN_PATH = r"C:/poppler-25.12.0/Library/bin" 

def is_mechanical_page(text):
    """
    Returns True if the page seems to be a Mechanical sheet.
    Checks for 'MECHANICAL', 'HVAC', or sheet numbers like 'M2.1'.
    """
    text = text.upper()
    if "MECHANICAL" in text or "HVAC" in text:
        return True
    # Regex for Sheet numbers (M-101, M2.0, etc.) often found in corners
    if re.search(r'\bM\s*[-.]?\s*\d', text):
        return True
    return False

def ocr_and_highlight_aggressive(pdf_bytes):
    """
    1. Converts to Images -> OCR.
    2. Scans word-by-word for patterns (AC-1).
    3. Highlights EVERYTHING matching the pattern.
    """
    try:
        # Higher DPI (250) for better reading of small text
        images = convert_from_bytes(pdf_bytes, dpi=250, poppler_path=POPPLER_BIN_PATH)
    except Exception as e:
        st.error(f"Poppler Error: {e}")
        return None, []

    # Create the Output PDF container
    pdf_writer = fitz.open()
    
    found_tags = []
    
    # Regex: Matches AC-1, AC 1, EF-10, etc.
    # Group 1: Type (AC), Group 3: Number (1)
    # Handles spaces/hyphens in between.
    pattern = re.compile(r'\b([A-Z]{1,4})(\s*[-‐‑_~]\s*|\s+)(\d+)\b', re.IGNORECASE)

    progress_bar = st.progress(0)
    status = st.empty()

    for i, img in enumerate(images):
        status.text(f"Processing Page {i+1}/{len(images)}...")
        progress_bar.progress((i + 1) / len(images))

        # 1. OCR to get a searchable PDF page
        try:
            pdf_page_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            page_doc = fitz.open("pdf", pdf_page_bytes)
            page = page_doc[0]  # There's only one page in this temp doc
        except Exception as e:
            st.warning(f"OCR failed on page {i+1}: {e}")
            continue

        # 2. Check if Mechanical Page (Optional Filter)
        full_text = page.get_text()
        if not is_mechanical_page(full_text):
            # If strictly required, skip. For now, we process ALL but flag it in Excel.
            is_mech = False
        else:
            is_mech = True

        # 3. Aggressive Word Scanning
        # We look for matches in the full text, then find their locations
        matches = pattern.finditer(full_text)
        
        # We use a set to prevent highlighting the same exact coordinate twice
        highlighted_rects = set()

        for match in matches:
            tag_str = match.group(0) # e.g. "AC - 1"
            tag_type = match.group(1).upper()
            tag_num = match.group(3)
            clean_tag = f"{tag_type}-{tag_num}"

            # Ask PyMuPDF: "Where exactly is this string?"
            # We search for the *raw OCR string* to ensure we find it.
            instances = page.search_for(tag_str)

            for inst in instances:
                # Coordinate Check (deduplication)
                # We round coordinates to avoid float precision issues
                rect_key = (round(inst.x0), round(inst.y0), round(inst.x1), round(inst.y1))
                if rect_key in highlighted_rects:
                    continue
                highlighted_rects.add(rect_key)

                # Color Coding
                # Cyan for Diffusers, Yellow for Units
                is_dist = tag_type in ['CD', 'RG', 'SR', 'SD']
                color = (0, 1, 1) if is_dist else (1, 1, 0)

                # HIGHLIGHT
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=color)
                annot.update()

                found_tags.append({
                    "Page": i + 1,
                    "Is_Mechanical_Sheet": is_mech,
                    "Tag": clean_tag,
                    "Raw_Text": tag_str,
                    "Category": "Distribution" if is_dist else "Equipment"
                })

        # Insert this processed page into our final PDF
        pdf_writer.insert_pdf(page_doc)

    progress_bar.empty()
    status.empty()

    out_buffer = io.BytesIO()
    pdf_writer.save(out_buffer)
    out_buffer.seek(0)
    
    return out_buffer.getvalue(), found_tags

# --- STREAMLIT UI ---
st.set_page_config(page_title="Deep Scan Highlighter", layout="wide")
st.title("🔍 Deep Scan & Highlight")
st.markdown("""
**Aggressive Mode:**
1.  **High-Res OCR (250 DPI):** Reads smaller text and blurry marks.
2.  **Pattern Match:** Highlights *anything* that looks like `Letter-Number` (e.g. `AC 1`, `EF-10`).
3.  **Coverage:** Highlights all instances found on the page.
""")

if "path" in POPPLER_BIN_PATH:
    st.error("⚠️ **Action Required:** Update `POPPLER_BIN_PATH` in the code!")
    st.stop()

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    raw_bytes = uploaded_file.read()

    with st.spinner("Running Deep Scan (OCR + Pattern Matching)..."):
        final_pdf_bytes, tags = ocr_and_highlight_aggressive(raw_bytes)

    if tags:
        df = pd.DataFrame(tags)
        
        # Summary
        st.success(f"Processing Complete! Highlighted {len(df)} items.")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            st.metric("Total Highlights", len(df))
            st.metric("Unique Tags", df['Tag'].nunique())
            
        with col2:
            st.subheader("Preview Found Tags")
            st.dataframe(df[['Page', 'Tag', 'Raw_Text', 'Category']].head(10), use_container_width=True)

        # Downloads
        c1, c2 = st.columns(2)
        
        c1.download_button(
            label="📥 Download Highlighted PDF",
            data=final_pdf_bytes,
            file_name="deep_scan_highlighted.pdf",
            mime="application/pdf"
        )
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        c2.download_button(
            label="📥 Download Tag List (Excel)",
            data=excel_buffer.getvalue(),
            file_name="tag_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("OCR completed but no tags were found. The scan quality might be too low, or the tags don't follow the 'Letter-Number' format.")