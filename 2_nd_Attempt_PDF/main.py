import tkinter as tk
import pandas
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
import os
import logging
from file_reader import read_excel_csv
from docx_filler import fill_docx_template
from pdf_generator import merge_pdfs
from data_mapper import map_data_to_docx
from docx2pdf import convert

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# Initialize GUI
root = tb.Window(themename="journal")
root.title("Automated Document Filler")
root.geometry("1000x600")

# Global Variables
input_file = None
template_file = None
output_folder = None

# 🟢 Load Excel/CSV File
def upload_data_file():
    global input_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.xls;*.csv")])
    if file_path:
        input_file = file_path
        lbl_data.config(text=f"📂 {os.path.basename(file_path)} Loaded")
        logging.info(f"Data file loaded: {file_path}")
        messagebox.showinfo("Success", "Data file loaded successfully!")

# 🟢 Load DOCX Template
def upload_template():
    global template_file
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        template_file = file_path
        lbl_template.config(text=f"📄 {os.path.basename(file_path)} Loaded")
        logging.info(f"Template file loaded: {file_path}")
        messagebox.showinfo("Success", "Template loaded successfully!")

# 🟢 Select Output Folder
def select_output_folder():
    global output_folder
    folder = filedialog.askdirectory()
    if folder:
        output_folder = folder
        lbl_output.config(text=f"📁 Output Folder: {folder}")
        logging.info(f"Output folder selected: {folder}")

# 🟢 Start Automated Processing
def start_processing():
    if not all([input_file, template_file, output_folder]):
        messagebox.showerror("Error", "Please upload all required files!")
        return

    try:
        data = read_excel_csv(input_file)
        if data is None:
            messagebox.showerror("Error", "Failed to read data file.")
            return

        generated_files = map_data_to_docx(
            template_path=template_file,
            data=data,
            output_folder=output_folder
        )

        if not generated_files:
            messagebox.showerror("Error", "No documents were generated.")
            return

        # Convert generated DOCX files to PDF
        pdf_output_folder = os.path.join(output_folder, "PDF_Output") #create pdf subfolder.
        os.makedirs(pdf_output_folder, exist_ok=True) #make sure subfolder exists.

        for docx_file in generated_files:
            pdf_file = os.path.join(pdf_output_folder, os.path.splitext(os.path.basename(docx_file))[0] + ".pdf")
            try:
                convert(docx_file, pdf_file)
                logging.info(f"Converted {docx_file} to {pdf_file}")
            except Exception as e:
                logging.error(f"Error converting {docx_file} to PDF: {e}")
                messagebox.showerror("Error", f"Error converting {docx_file} to PDF: {e}")

        messagebox.showinfo("Success", "Documents generated and converted to PDF successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed: {str(e)}")
        logging.error(f"Processing error: {str(e)}")

# -------------------- DOCX to PDF Conversion Functions --------------------
def convert_docx_to_pdf(docx_folder, output_folder):
    """
    Converts all DOCX files in the given folder to PDF files,
    and saves them in the output folder.

    Args:
        docx_folder: Folder containing DOCX files.
        output_folder: Folder to save the generated PDF files.
    """
    if not os.path.exists(docx_folder):
        logging.error(f"DOCX folder not found: {docx_folder}")
        messagebox.showerror("Error", f"DOCX folder not found: {docx_folder}")
        return

    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            logging.error(f"Error creating output folder: {output_folder}, {e}")
            messagebox.showerror("Error", f"Error creating output folder: {output_folder}")
            return

    docx_files = [f for f in os.listdir(docx_folder) if f.endswith(".docx")]

    if not docx_files:
        logging.warning("No DOCX files found in the folder.")
        messagebox.showwarning("Warning", "No DOCX files found in the folder.")
        return

    for docx_file in docx_files:
        docx_path = os.path.join(docx_folder, docx_file)
        pdf_file = os.path.splitext(docx_file)[0] + ".pdf"
        pdf_path = os.path.join(output_folder, pdf_file)

        try:
            convert(docx_path, pdf_path)
            logging.info(f"Converted: {docx_file} to {pdf_file}")
        except Exception as e:
            logging.error(f"Error converting {docx_file}: {e}")
            messagebox.showerror("Error", f"Error converting {docx_file}: {e}")

def select_folders_and_convert():
    """Selects input and output folders using file dialogs and starts conversion."""
    docx_folder = filedialog.askdirectory(title="Select DOCX Folder")
    if not docx_folder:
        return  # User cancelled

    output_folder = filedialog.askdirectory(title="Select Output PDF Folder")
    if not output_folder:
        return  # User cancelled

    convert_docx_to_pdf(docx_folder, output_folder)
    messagebox.showinfo("Conversion Complete", "DOCX to PDF conversion completed.")
# --------------------------------------------------------------------------

# 🔹 GUI Layout
frame = tb.Frame(root)
frame.pack(pady=20)

btn_data = tb.Button(frame, text="📂 Upload Data File", command=upload_data_file)
btn_data.grid(row=0, column=0, padx=10, pady=5)

btn_template = tb.Button(frame, text="📄 Upload DOCX Template", command=upload_template)
btn_template.grid(row=1, column=0, padx=10, pady=5)

btn_output = tb.Button(frame, text="📁 Select Output Folder", command=select_output_folder)
btn_output.grid(row=2, column=0, padx=10, pady=5)

btn_start = tb.Button(frame, text="🚀 Start Processing", bootstyle="success", command=start_processing)
btn_start.grid(row=3, column=0, padx=10, pady=20)

btn_docx_to_pdf = tb.Button(frame, text="📄 to 📄 Convert DOCX to PDF", bootstyle="info", command=select_folders_and_convert)
btn_docx_to_pdf.grid(row=4, column=0, padx=10, pady=20) #add the new button

# Labels for file paths
lbl_data = tb.Label(frame, text="No Data File Loaded", bootstyle="secondary")
lbl_data.grid(row=0, column=1, padx=10, sticky="w")

lbl_template = tb.Label(frame, text="No Template File Loaded", bootstyle="secondary")
lbl_template.grid(row=1, column=1, padx=10, sticky="w")

lbl_output = tb.Label(frame, text="No Output Folder Selected", bootstyle="secondary")
lbl_output.grid(row=2, column=1, padx=10, sticky="w")

root.mainloop()