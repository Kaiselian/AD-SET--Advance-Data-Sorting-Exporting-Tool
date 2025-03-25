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

        # Process all rows into separate documents
        generated_files = map_data_to_docx(
            template_path=template_file,
            data=data,  # Your pandas DataFrame
            output_folder=output_folder
        )

        if not generated_files:
            messagebox.showerror("Error", "No invoices were generated")
        else:
            messagebox.showinfo("Success",
                                f"Generated {len(generated_files)} invoices:\n" +
                                "\n".join(os.path.basename(f) for f in generated_files[:3]) +
                                ("\n..." if len(generated_files) > 3 else "")
                                )

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed: {str(e)}")
        logging.error(f"Processing error: {str(e)}")

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

# Labels for file paths
lbl_data = tb.Label(frame, text="No Data File Loaded", bootstyle="secondary")
lbl_data.grid(row=0, column=1, padx=10, sticky="w")

lbl_template = tb.Label(frame, text="No Template File Loaded", bootstyle="secondary")
lbl_template.grid(row=1, column=1, padx=10, sticky="w")

lbl_output = tb.Label(frame, text="No Output Folder Selected", bootstyle="secondary")
lbl_output.grid(row=2, column=1, padx=10, sticky="w")

root.mainloop()
