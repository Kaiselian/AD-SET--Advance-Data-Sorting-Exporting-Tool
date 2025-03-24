import fitz
from PIL import Image
from PyQt5.QtGui import QPixmap, QImage

def add_text_to_pdf(pdf_path, boxes_and_columns, data, output_path):
    doc = fitz.open(pdf_path)
    page = doc[0]  # Assuming first page

    for box, column in boxes_and_columns.items():
        text = str(data.iloc[0][column])  # Get data from the first row
        x = box.x()
        y = box.y()

        page.insert_text((x, y), text)

    doc.save(output_path)

def load_pdf(pdf_path):
    """
    Loads the first page of a PDF as a QPixmap.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        tuple: (QPixmap, fitz.Document) or (None, None) in case of error.
    """
    try:
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[0]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((800, 1000))  # Resize the image

        # Convert PIL Image to QImage
        qimage = QImage(img.tobytes("raw", "RGB"), img.width, img.height, QImage.Format_RGB888)

        # Convert QImage to QPixmap
        pdf_pixmap = QPixmap.fromImage(qimage)

        return pdf_pixmap, pdf_document
    except Exception as e:
        print(f"Error loading PDF: {e}")
        return None, None