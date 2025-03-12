from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError
from docx2pdf import convert
from reportlab.pdfgen import canvas
import os

def generate_pdfs(docx_files, output_folder):
    """
    Converts filled DOCX files to PDFs.

    :param docx_files: List of DOCX files to convert
    :param output_folder: Folder where PDFs will be saved
    :return: List of generated PDF file paths
    """
    pdf_files = []

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for docx_file in docx_files:
        pdf_output = os.path.join(output_folder, os.path.basename(docx_file).replace(".docx", ".pdf"))

        # Skip if PDF already exists
        if os.path.exists(pdf_output):
            print(f"⚠️ PDF already exists: {pdf_output}, skipping conversion.")
        else:
            try:
                convert(docx_file, pdf_output)
                print(f"✅ Converted: {docx_file} → {pdf_output}")
                pdf_files.append(pdf_output)
            except Exception as e:
                print(f"❌ Error converting {docx_file} to PDF: {e}")

    return pdf_files

def merge_pdfs(input_folder, output_pdf):
    """
    Merges all PDFs in the given folder into a single PDF.

    :param input_folder: Folder containing individual PDFs
    :param output_pdf: Path to save the merged PDF
    """
    pdf_writer = PdfWriter()
    pdf_files = sorted([
        os.path.join(input_folder, f) for f in os.listdir(input_folder)
        if f.endswith(".pdf") and os.path.isfile(os.path.join(input_folder, f))
    ])

    if not pdf_files:
        print("❌ No PDFs found in the folder. Cannot merge.")
        return

    for pdf_file in pdf_files:
        try:
            pdf_reader = PdfReader(pdf_file)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
        except FileNotFoundError:
            print(f"❌ Error: PDF file not found: {pdf_file}")
        except PdfReadError:
            print(f"❌ Error: Could not read PDF file: {pdf_file}. It might be corrupted.")
        except Exception as e:
            print(f"❌ Error merging {pdf_file}: {e}")

    try:
        with open(output_pdf, "wb") as output:
            pdf_writer.write(output)
        print(f"✅ Merged PDF saved: {output_pdf}")
    except Exception as e:
        print(f"❌ Error saving merged PDF: {e}")
