from PyPDF2 import PdfReader, PdfWriter
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def merge_pdfs(input_folder, output_pdf, sort_key=None):
    """
    Merges all PDFs in the given folder into a single PDF.

    :param input_folder: Folder containing individual PDFs
    :param output_pdf: Path to save the merged PDF
    :param sort_key: Optional function to sort PDF filenames (e.g., lambda x: int(x.split('_')[1]))
    """
    pdf_writer = PdfWriter()
    pdf_files = [f for f in os.listdir(input_folder) if f.endswith(".pdf")]

    if not pdf_files:
        logging.error("❌ No PDFs found in the folder.")
        return

    # Sort PDF files
    if sort_key:
        pdf_files.sort(key=sort_key)
    else:
        pdf_files.sort()

    logging.info(f"Found {len(pdf_files)} PDFs to merge.")

    for i, pdf_file in enumerate(pdf_files, start=1):
        pdf_path = os.path.join(input_folder, pdf_file)
        try:
            pdf_reader = PdfReader(pdf_path)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
            logging.info(f"✅ Added {pdf_file} ({i} of {len(pdf_files)})")
        except Exception as e:
            logging.error(f"❌ Error reading {pdf_file}: {e}")

    try:
        with open(output_pdf, "wb") as output:
            pdf_writer.write(output)
        logging.info(f"✅ Merged PDF saved: {output_pdf}")
    except Exception as e:
        logging.error(f"❌ Error saving merged PDF: {e}")

# Example usage
if __name__ == "__main__":
    input_folder = "C:/Users/anich/Downloads/Output"  # Folder containing individual PDFs
    output_pdf = "C:/Users/anich/Downloads/Merged.pdf"  # Path to save the merged PDF
    merge_pdfs(input_folder, output_pdf)