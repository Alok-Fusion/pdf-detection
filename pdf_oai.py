import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import os

# 👉 Uncomment and set path ONLY if Windows
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def make_pdf_searchable(input_pdf, output_pdf, dpi=300):
    """
    Converts any scanned PDF into a fully searchable & selectable PDF.
    """

    input_doc = fitz.open(input_pdf)
    output_doc = fitz.open()

    for page_num in range(len(input_doc)):
        page = input_doc[page_num]

        # Render page to image
        pix = page.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))

        # OCR → PDF page with invisible text layer
        ocr_pdf_bytes = pytesseract.image_to_pdf_or_hocr(
            img, extension="pdf"
        )

        ocr_doc = fitz.open(stream=ocr_pdf_bytes, filetype="pdf")
        output_doc.insert_pdf(ocr_doc)

        print(f"✔ OCR completed for page {page_num + 1}")

    output_doc.save(output_pdf)
    output_doc.close()
    input_doc.close()

    print("\n🎉 DONE! Your PDF is now fully searchable & selectable.")


if __name__ == "__main__":
    INPUT_PDF = "C:/Users/ak500/Downloads/mechanical (2).pdf"
    OUTPUT_PDF = "searchable_output.pdf"

    make_pdf_searchable(INPUT_PDF, OUTPUT_PDF)
