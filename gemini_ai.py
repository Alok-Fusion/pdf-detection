import streamlit as st
import fitz
import pytesseract
from pdf2image import convert_from_bytes
import io
import json
import requests
import base64

# --- CONFIGURATION ---
POPPLER_BIN_PATH = r"C:/Alok/poppler-25.12.0/Library/bin"

def call_gemini_direct(image_bytes, api_key):
    """
    Calls Gemini API directly via HTTP (No library version issues).
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    
    # Encode image
    b64_img = base64.b64encode(image_bytes).decode('utf-8')
    
    payload = {
        "contents": [{
            "parts": [
                {"text": "Find mechanical tags (e.g. AC-1, EF-5). Return ONLY JSON: {\"tags\": [\"AC-1\"]}"},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64_img}}
            ]
        }]
    }
    
    try:
        response = requests.post(url, json=payload, headers={"Content-Type": "application/json"})
        if response.status_code != 200:
            return []
            
        result = response.json()
        text_resp = result['candidates'][0]['content']['parts'][0]['text']
        
        # Clean JSON
        clean_json = text_resp.replace("```json", "").replace("```", "").strip()
        data = json.loads(clean_json)
        return data.get("tags", [])
    except:
        return []

def run_ai_pipeline(pdf_bytes, api_key):
    images = convert_from_bytes(pdf_bytes, dpi=200, poppler_path=POPPLER_BIN_PATH)
    pdf_writer = fitz.open()
    
    st_bar = st.progress(0)
    
    for i, img in enumerate(images):
        st_bar.progress((i+1)/len(images))
        
        # 1. Convert Img to Bytes
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        
        # 2. Get Tags from AI
        tags = call_gemini_direct(img_byte_arr.getvalue(), api_key)
        
        # 3. Create Searchable Page
        pdf_page = pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
        page_doc = fitz.open("pdf", pdf_page)
        page = page_doc[0]
        
        # 4. Highlight
        for tag in tags:
            # Loose search (AC-1 -> AC 1)
            variants = [tag, tag.replace("-", " "), tag.replace("-", "")]
            for v in variants:
                insts = page.search_for(v)
                for inst in insts:
                    page.add_highlight_annot(inst).update()
                    
        pdf_writer.insert_pdf(page_doc)
        
    out = io.BytesIO()
    pdf_writer.save(out)
    return out.getvalue()

# --- UI ---
st.title("Gemini Fix (Direct API)")
key = st.text_input("API Key", type="password")
f = st.file_uploader("PDF", type="pdf")
if f and key:
    if "path" in POPPLER_BIN_PATH: st.error("Fix Poppler Path"); st.stop()
    res = run_ai_pipeline(f.read(), key)
    st.download_button("Download", res, "ai_fixed.pdf")