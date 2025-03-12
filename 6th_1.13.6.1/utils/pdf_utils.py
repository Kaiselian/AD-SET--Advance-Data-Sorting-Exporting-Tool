import fitz
from PIL import Image, ImageTk

def load_pdf(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[0]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((800, 1000))
        pdf_img = ImageTk.PhotoImage(img)
        return pdf_img, pdf_document
    except Exception as e:
        print(f"Error loading PDF: {e}")
        return None, None