from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF for better PDF handling
import os

def replace_pdf_placeholders(input_pdf, output_pdf, data):
    """Replaces placeholders in a PDF with data, using PyMuPDF and ReportLab."""
    try:
        # Read the PDF
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        doc = fitz.open(input_pdf)

        # Create a temporary PDF with ReportLab
        temp_pdf = "temp_overlay.pdf"
        c = canvas.Canvas(temp_pdf, pagesize=letter)
        c.setFont("Helvetica", 12)

        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                for placeholder, value in data.items():
                    if placeholder in text:
                        print(f"✅ Found placeholder: {placeholder} on page {page_num + 1}")
                        x_start = 1 * inch  # Adjusted X position
                        y_start = 9 * inch - (page_num * 0.5 * inch)  # Adjusted Y position
                        box_width = 6 * inch
                        box_height = 1 * inch

                        draw_text_in_box(c, str(value), x_start, y_start, box_width, box_height)

        c.save()

        # Merge overlay onto original PDF
        overlay = fitz.open(temp_pdf)
        for page_num in range(len(doc)):
            doc[page_num].insert_pdf(overlay, from_page=page_num, to_page=page_num)

        # Save final output
        doc.save(output_pdf)
        doc.close()
        overlay.close()
        os.remove(temp_pdf)

        print(f"✅ Modified PDF saved: {output_pdf}")

    except FileNotFoundError:
        print("❌ Error: PDF not found.")
    except Exception as e:
        print(f"❌ Error: {e}")

def draw_text_in_box(canvas, text, x, y, width, height, font_name="Helvetica", font_size=12):
    """Draws text inside a specified box with automatic word wrapping."""
    canvas.setFont(font_name, font_size)
    lines = []
    words = text.split()
    current_line = ""

    for word in words:
        test_line = current_line + " " + word if current_line else word
        text_width = canvas.stringWidth(test_line, font_name, font_size)
        if text_width <= width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word

    lines.append(current_line)

    line_height = font_size * 1.2
    for i, line in enumerate(lines):
        if (i + 1) * line_height <= height:
            canvas.drawString(x, y - i * line_height, line)
        else:
            canvas.drawString(x, y - i * line_height, line + "...")  # Show ellipsis if text is too long
            break

# Example Usage
if __name__ == "__main__":
    input_pdf = "LensKart Invoice-Template.pdf"
    output_pdf = "output_invoice.pdf"
    data = {
        "{{Shipment Code:}}": "SHIP12345",
        "{{Order:}}": "ORDER67890",
        "{{Total Amount:}}": "$150.00",
    }

    replace_pdf_placeholders(input_pdf, output_pdf, data)
