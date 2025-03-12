from docx2pdf import convert
import os
import tkinter as tk
from tkinter import messagebox

def generate_pdfs(docx_files, output_folder):
    """Converts a list of DOCX files to PDFs and saves them in the output folder."""
    pdf_files = []
    for docx_file in docx_files:
        if not os.path.exists(docx_file):
            print(f"ERROR: DOCX file not found: {docx_file}")
            continue

        pdf_file = os.path.join(output_folder, os.path.splitext(os.path.basename(docx_file))[0] + ".pdf")
        try:
            convert(docx_file, pdf_file)
            pdf_files.append(pdf_file)
        except Exception as e:
            print(f"ERROR: Failed to convert {docx_file}: {e}")

    if pdf_files:
        print(f"INFO: Successfully converted {len(pdf_files)} files to PDF.")
        return pdf_files
    else:
        print("INFO: Successfully converted 0 files to PDF.")
        return None