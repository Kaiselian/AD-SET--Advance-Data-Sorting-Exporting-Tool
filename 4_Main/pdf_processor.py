import os
import logging
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from tempfile import NamedTemporaryFile

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

def extract_placeholders_from_pdf(pdf_path):
    """
    Extracts placeholders from a PDF.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        set: Set of placeholders, or None if an error occurs.
    """
    try:
        reader = PdfReader(pdf_path)
        placeholders = set()

        for page in reader.pages:
            text = page.extract_text()
            if text:
                detected = {word for word in text.split() if word.startswith("{{") and word.endswith("}}")}
                placeholders.update(detected)

        logger.info(f"Extracted {len(placeholders)} placeholders from {pdf_path}")
        return placeholders

    except Exception as e:
        logger.error(f"Error reading PDF: {e}")
        return None

def replace_pdf_placeholders(input_pdf, output_pdf, data, font="Helvetica", font_size=12):
    """
    Replaces placeholders in a PDF using an overlay technique.

    Args:
        input_pdf (str): Path to the input PDF file.
        output_pdf (str): Path to save the output PDF file.
        data (dict): Dictionary mapping placeholders to replacement values.
        font (str): Font name for the replacement text (default: "Helvetica").
        font_size (int): Font size for the replacement text (default: 12).
    """
    try:
        # Validate inputs
        if not os.path.exists(input_pdf):
            logger.error(f"Input PDF file not found: {input_pdf}")
            return
        if not data:
            logger.error("No data provided for replacement.")
            return

        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        # Create a temporary overlay PDF
        with NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
            temp_path = temp_file.name
            c = canvas.Canvas(temp_path, pagesize=letter)
            c.setFont(font, font_size)

            # Replace placeholders on the first page
            page = reader.pages[0]
            text = page.extract_text()
            if text:
                for placeholder, value in data.items():
                    if placeholder in text:
                        # Example: Place text at fixed coordinates (customize as needed)
                        c.drawString(100, 700, str(value))

            c.save()

            # Merge the overlay with the original PDF
            overlay_reader = PdfReader(temp_path)
            overlay_page = overlay_reader.pages[0]
            page.merge_page(overlay_page)
            writer.add_page(page)

            # Save the output PDF
            with open(output_pdf, "wb") as f:
                writer.write(f)

            logger.info(f"PDF saved: {output_pdf}")

        # Clean up temporary file
        os.remove(temp_path)

    except Exception as e:
        logger.error(f"Error processing PDF: {e}")