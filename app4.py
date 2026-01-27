import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import fitz  # PyMuPDF
import pandas as pd
import io
import re

# --- CONFIGURATION ---
# Update this to your Poppler bin path
POPPLER_BIN_PATH = r"C:/Alok/poppler-25.12.0/Library/bin"

def is_mechanical_page(text):
    """
    Simple check to see if the page is likely a Mechanical drawing.
    Looks for 'Sheet M', 'M-', 'Mechanical', or 'HVAC' in the text.
    """
    text_upper = text.upper()
    # Check for common indicators found in title blocks
    if "MECHANICAL" in text_upper or "HVAC" in text_upper:
        return True
    # Check for sheet numbers like M2.1, M-101 at the end of the text
    # (Title blocks are usually at the end of the text stream)
    if re.search(r'\bM[-.]?\d', text_upper):
        return True
    return False

def create_searchable_pdf(images):
    """
    Converts images to a single multi-page PDF with hidden text layer.
    """
    pdf_writer = fitz.open()
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, img in enumerate(images):
        status_text.text(f"OCR Phase: Reading Page {i+1}...")
        progress_bar.progress((i + 1) / len(images))
        try:
            # Tesseract creates a PDF page with text
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
            img_pdf = fitz.open("pdf", pdf_bytes)
            pdf_writer.insert_pdf(img_pdf)
        except Exception as e:
            st.error(f"OCR Error on page {i+1}: {e}")
            
    progress_bar.empty()
    status_text.empty()
    return pdf_writer

def highlight_all_tags(doc):
    """
    Iterates through EVERY word on the page.
    If a word matches the Tag Pattern, it gets highlighted.
    """
    extracted_data = []
    
    # Regex: Matches AC-1, AC-12, EF-5, etc.
    # We use ^ and $ to ensure we match the WHOLE word (avoiding partials)
    tag_pattern = re.compile(r'^(AC|CU|EF|HP|AH|CD|RG|SR|SD)[-_]?(\d+)$', re.IGNORECASE)

    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # 1. Get all words with coordinates
        # Returns list of: (x0, y0, x1, y1, "word", block_no, line_no, word_no)
        words = page.get_text("words")
        
        # 2. Check Page Content for "Mechanical" filtering
        # We reconstruct the page text roughly to check for "Sheet M..."
        full_text = " ".join([w[4] for w in words])
        
        if not is_mechanical_page(full_text):
            # If you want to force ALL pages, verify this logic or remove this block
            # For now, we note it but don't skip entirely (safer), or mark it in data
            is_mech = False
        else:
            is_mech = True

        # 3. Iterate EVERY word
        for w in words:
            text = w[4].strip()
            
            # Clean punctuation (sometimes OCR adds a dot like "AC-1.")
            text_clean = text.strip(".,")
            
            match = tag_pattern.match(text_clean)
            if match:
                # We found a tag!
                tag_type = match.group(1).upper()
                tag_num = match.group(2)
                full_tag = f"{tag_type}-{tag_num}"
                
                # Coordinates from the word list
                rect = fitz.Rect(w[0], w[1], w[2], w[3])
                
                # Color Coding
                is_dist = tag_type in ['CD', 'RG', 'SR', 'SD']
                color = (0, 1, 1) if is_dist else (1, 1, 0) # Cyan or Yellow

                # Highlight
                annot = page.add_highlight_annot(rect)
                annot.set_colors(stroke=color)
                annot.update()

                extracted_data.append({
                    "Page": page_num + 1,
                    "Tag": full_tag,
                    "Is_Mechanical_Sheet": is_mech,
                    "X": round(w[0], 2),
                    "Y": round(w[1], 2)
                })

    return doc, extracted_data

# --- STREAMLIT UI ---
st.set_page_config(page_title="Complete Tag Highlighter", layout="wide")
st.title("🔎 Complete Mechanical Tag Highlighter")
st.markdown("""
**New Logic:**
1.  **OCR:** Makes the PDF readable.
2.  **Word-by-Word Search:** Checks every single word. If it sees `AC-1`, it highlights it.
3.  **Completeness:** Highlights *every* occurrence, not just the first one.
""")

uploaded_file = st.file_uploader("Upload Scanned PDF", type="pdf")

if uploaded_file:
    # Check Config
    if "path" in POPPLER_BIN_PATH:
        st.error("⚠️ Update POPPLER_BIN_PATH in the code code!")
        st.stop()

    # 1. Conversion
    with st.spinner("Step 1/3: Reading Scan..."):
        try:
            images = convert_from_bytes(uploaded_file.read(), dpi=200, poppler_path=POPPLER_BIN_PATH)
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    # 2. OCR (Make Searchable)
    with st.spinner("Step 2/3: Recognizing Text..."):
        doc = create_searchable_pdf(images)
    
    # 3. Highlight
    with st.spinner("Step 3/3: Highlighting Every Tag..."):
        highlighted_doc, data = highlight_all_tags(doc)

    # 4. Results
    if data:
        df = pd.DataFrame(data)
        
        # Filter for display (Show counts)
        st.success(f"Processing Complete! Found {len(df)} total tags.")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.metric("Total Highlights", len(df))
            st.metric("Mechanical Sheets Detected", df['Is_Mechanical_Sheet'].sum() if not df.empty else 0)
        
        with col2:
            st.write("Preview of Found Tags:")
            st.dataframe(df.head(10), use_container_width=True)

        # Downloads
        st.divider()
        d1, d2 = st.columns(2)
        
        # Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        d1.download_button("📥 Download Excel Log", excel_buffer.getvalue(), "all_tags.xlsx")
        
        # PDF
        out_pdf_buffer = io.BytesIO()
        highlighted_doc.save(out_pdf_buffer)
        out_pdf_buffer.seek(0)
        d2.download_button("🖍️ Download Fully Highlighted PDF", out_pdf_buffer, "marked_complete.pdf", "application/pdf")
        
    else:
        st.warning("OCR finished but no tags like 'AC-1' were found.")