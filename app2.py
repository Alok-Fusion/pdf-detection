import streamlit as st
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes
import fitz  # PyMuPDF
import pandas as pd
import io
import re

# --- CONFIGURATION ---
# 1. Update this to your Poppler bin path
POPPLER_BIN_PATH = r"C:/poppler-25.12.0/Library/bin" 

def ocr_and_mark(pdf_file):
    """
    1. Converts PDF to Images.
    2. OCRs the text.
    3. Draws a VISIBLE RED BOX annotation on top of the PDF.
    """
    # Load PDF
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    pdf_file.seek(0)
    
    # Convert to Images (200 DPI is good balance)
    try:
        images = convert_from_bytes(pdf_file.read(), dpi=200, poppler_path=POPPLER_BIN_PATH)
    except Exception as e:
        st.error(f"Poppler Error: {e}")
        return None, []

    extracted_data = []
    
    # Regex: Matches AC-1, AC 1, AC - 1
    tag_pattern = re.compile(r'(AC|CU|EF|HP|AH|CD|RG|SR|SD)\s*[-_ ]\s*(\d+)', re.IGNORECASE)

    progress_bar = st.progress(0)
    status_text = st.empty()

    for page_num, img in enumerate(images):
        status_text.text(f"Processing Page {page_num + 1}...")
        progress_bar.progress((page_num + 1) / len(images))

        # 1. Run OCR
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT)
        n_boxes = len(ocr_data['text'])
        
        # 2. Coordinate Math
        # We need to map Image Pixels (from OCR) -> PDF Points (for drawing)
        fitz_page = doc[page_num]
        
        # Get dimensions
        pdf_w, pdf_h = fitz_page.rect.width, fitz_page.rect.height
        img_w, img_h = img.width, img.height
        
        scale_x = pdf_w / img_w
        scale_y = pdf_h / img_h

        # 3. Find & Mark Tags
        for i in range(n_boxes):
            text = ocr_data['text'][i].strip()
            if not text: continue

            # Check Match
            match = tag_pattern.match(text)
            if match:
                full_tag = f"{match.group(1).upper()}-{match.group(2)}"
                
                # OCR Coordinates (Pixels)
                x, y, w, h = (ocr_data['left'][i], ocr_data['top'][i], ocr_data['width'][i], ocr_data['height'][i])
                
                # Convert to PDF Coordinates
                # We add a little 'padding' (-2/+4) to make the box frame the text nicely
                rect_x0 = (x - 2) * scale_x
                rect_y0 = (y - 2) * scale_y
                rect_x1 = (x + w + 4) * scale_x
                rect_y1 = (y + h + 4) * scale_y
                
                # --- THE FIX: Use add_rect_annot (Sticker on top) ---
                rect = fitz.Rect(rect_x0, rect_y0, rect_x1, rect_y1)
                
                # Determine Color (Red for Equipment, Cyan for Diffusers)
                is_dist = full_tag[:2] in ['CD', 'RG', 'SR', 'SD']
                color = (0, 1, 1) if is_dist else (1, 0, 0) # Cyan or Red
                
                annot = fitz_page.add_rect_annot(rect)
                annot.set_border(width=2)
                annot.set_colors(stroke=color) # Border color
                annot.update()

                extracted_data.append({
                    "Page": page_num + 1,
                    "Tag": full_tag,
                    "Detected Text": text,
                    "X": round(rect_x0, 2),
                    "Y": round(rect_y0, 2)
                })

    progress_bar.empty()
    status_text.empty()
    
    # Save Result
    out_pdf = io.BytesIO()
    doc.save(out_pdf)
    out_pdf.seek(0)
    
    return out_pdf, extracted_data

# --- APP UI ---
st.set_page_config(page_title="OCR Tag Scanner")
st.title("📷 OCR Mechanical Scanner (Visible Boxes)")

uploaded_file = st.file_uploader("Upload Scanned PDF", type="pdf")

if uploaded_file:
    # --- CHECK POPPLER CONFIG ---
    if "path" in POPPLER_BIN_PATH:
        st.error("⚠️ You forgot to update the POPPLER_BIN_PATH in the code code!")
        st.stop()
        
    pdf_out, data = ocr_and_mark(uploaded_file)
    
    if data:
        st.success(f"✅ Found and Marked {len(data)} tags.")
        
        # Display Data
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)
        
        # Downloads
        col1, col2 = st.columns(2)
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        col1.download_button("📥 Download Excel", excel_buffer.getvalue(), "tags.xlsx")
        col2.download_button("🖍️ Download Marked PDF", pdf_out, "marked_plans.pdf", "application/pdf")
    else:
        st.warning("OCR finished but found no tags matching 'AC-X', 'EF-X', etc.")