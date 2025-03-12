from PyPDF2 import PdfReader, PdfWriter
import os

def generate_pdfs(input_docx, output_folder):
    # Your function logic here
    pass  # Placeholder

def merge_pdfs(input_folder, output_pdf):
    """
    Merges all PDFs in the given folder into a single PDF.

    :param input_folder: Folder containing individual PDFs
    :param output_pdf: Path to save the merged PDF
    """
    pdf_writer = PdfWriter()
    pdf_files = sorted([f for f in os.listdir(input_folder) if f.endswith(".pdf")])

    if not pdf_files:
        print("❌ No PDFs found in the folder.")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        pdf_reader = PdfReader(pdf_path)

        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

    with open(output_pdf, "wb") as output:
        pdf_writer.write(output)

    print(f"✅ Merged PDF saved: {output_pdf}")

