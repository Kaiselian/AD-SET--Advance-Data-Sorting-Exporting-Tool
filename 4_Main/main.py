import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
import os
from file_reader import read_excel_csv
from data_mapper import map_data_to_docx
from docx_filler import fill_docx_template
from pdf_generator import generate_pdfs
from pdf_processor import extract_placeholders_from_pdf, replace_pdf_placeholders
from data_viewer import view_data_table


# Global Variables
input_file = None
template_file = None
pdf_template_file = None  # Added this variable for PDFs
output_folder = None
data = None

def upload_data_file():
    global input_file  # Declare global variable
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.xls;*.csv")])
    if file_path:
        input_file = file_path
        lbl_data.config(text=f"📂 {os.path.basename(file_path)} Loaded", bootstyle="success")
        messagebox.showinfo("Success", "Data file loaded successfully!")

def upload_template():
    global template_file  # Declare global variable
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        template_file = file_path
        lbl_template.config(text=f"📄 {os.path.basename(file_path)} Loaded", bootstyle="success")
        messagebox.showinfo("Success", "Template loaded successfully!")

def select_output_folder():
    global output_folder  # Declare global variable
    folder = filedialog.askdirectory()
    if folder:
        output_folder = folder
        lbl_output.config(text=f"📁 Output Folder: {folder}", bootstyle="success")

# 🟢 Load PDF Template
def upload_pdf_template():
    global pdf_template_file  # Declare global variable
    file_path = filedialog.askopenfilename(filetypes=[("PDF Documents", "*.pdf")])
    if file_path:
        pdf_template_file = file_path
        lbl_pdf_template.config(text=f"📄 {os.path.basename(file_path)} Loaded", bootstyle="success")
        messagebox.showinfo("Success", "PDF Template loaded successfully!")

        # Extract placeholders and display them
        placeholders = extract_placeholders_from_pdf(pdf_template_file)
        if placeholders:
            placeholder_text.set("\n".join(placeholders))
        else:
            placeholder_text.set("⚠️ No placeholders found!")

        messagebox.showinfo("Success", "PDF Template loaded successfully!")
    else:
        lbl_pdf_template.config(text="⚠️ No PDF Template Loaded", bootstyle="danger")

# 🟢 Start Processing
def start_processing():
    if not input_file or (not template_file and not pdf_template_file) or not output_folder:
        messagebox.showerror("Error", "Please upload all required files!")
        return

    # Progress Bar Setup
    progress_window = tk.Toplevel(root)
    progress_window.title("Processing...")
    progress_bar = tb.Progressbar(progress_window, mode="indeterminate")
    progress_bar.pack(pady=20, padx=20)
    progress_bar.start()

    try:
        # Step 1: Read Data
        data = read_excel_csv(input_file)
        if data is None:
            messagebox.showerror("Error", "Failed to read data file.")
            return

        if template_file:  # DOCX Processing
            # Step 2: Map Data to DOCX Placeholders
            mapped_data = map_data_to_docx(template_file, data)
            if mapped_data is None:
                messagebox.showerror("Error", "Failed to map data to template.")
                return

            # Step 3: Fill DOCX Template
            filled_files = fill_docx_template(template_file, mapped_data, output_folder)
            if filled_files is None:
                messagebox.showerror("Error", "Failed to fill DOCX template.")
                return

            # Step 4: Convert to PDFs
            generate_pdfs(filled_files, output_folder)
            messagebox.showinfo("Success", "All PDFs generated successfully!")

        elif pdf_template_file:  # PDF Processing
            # Step 2: Map Data to PDF Placeholders
            pdf_data = {f"{{{{{col.strip()}}}}}": str(data[col].iloc[0]) for col in data.columns}
            output_pdf_file = os.path.join(output_folder, "output_invoice.pdf")
            replace_pdf_placeholders(pdf_template_file, output_pdf_file, pdf_data)
            messagebox.showinfo("Success", "PDF generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        progress_bar.stop()
        progress_window.destroy()

#data_viewer
def view_data():
    """Opens the data viewer to display the uploaded data."""
    global data
    if data is None or data.empty:
        messagebox.showerror("Error", "No data file uploaded!")
        return
    view_data_table(root, data)

# 🔹 GUI Setup
root = tb.Window(themename="journal")
root.title("Automated Document Filler")
root.geometry("1000x600")

frame = tb.Frame(root)
frame.pack(pady=20)

btn_data = tb.Button(frame, text="📂 Upload Data File", command=upload_data_file)
btn_data.grid(row=0, column=0, padx=10, pady=5)

btn_view_data = tb.Button(frame, text="📊 View Data", bootstyle="info", command=view_data)
btn_view_data.grid(row=0, column=1, padx=10, pady=5)

btn_template = tb.Button(frame, text="📄 Upload DOCX Template", command=upload_template)
btn_template.grid(row=1, column=0, padx=10, pady=5)

btn_pdf_template = tb.Button(frame, text="📄 Upload PDF Template", command=upload_pdf_template) #add pdf template button.
btn_pdf_template.grid(row=2, column=0, padx=10, pady=5)

btn_output = tb.Button(frame, text="📁 Select Output Folder", command=select_output_folder)
btn_output.grid(row=3, column=0, padx=10, pady=5)

btn_start = tb.Button(frame, text="🚀 Start Processing", bootstyle="success", command=start_processing)
btn_start.grid(row=4, column=0, padx=10, pady=20)


# 🟢 Labels for file paths
lbl_data = tb.Label(frame, text="No Data File Loaded", bootstyle="secondary")
lbl_data.grid(row=0, column=1, padx=10, sticky="w")

lbl_template = tb.Label(frame, text="No Template File Loaded", bootstyle="secondary")
lbl_template.grid(row=1, column=1, padx=10, sticky="w")

lbl_pdf_template = tb.Label(frame, text="No PDF Template File Loaded", bootstyle="secondary")
lbl_pdf_template.grid(row=2, column=1, padx=10, sticky="w")

lbl_output = tb.Label(frame, text="No Output Folder Selected", bootstyle="secondary")
lbl_output.grid(row=3, column=1, padx=10, sticky="w")
btn_pdf_template = tb.Button(frame, text="📄 Upload PDF Template", command=upload_pdf_template)
btn_pdf_template.grid(row=2, column=0, padx=10, pady=5)

btn_start = tb.Button(frame, text="🚀 Start Processing", bootstyle="success", command=start_processing)
btn_start.grid(row=4, column=0, padx=10, pady=20)

placeholder_text = tk.StringVar(value="No placeholders detected")
lbl_placeholders = tb.Label(frame, textvariable=placeholder_text, bootstyle="info", justify="left")
lbl_placeholders.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

lbl_pdf_template = tb.Label(frame, text="No PDF Template Loaded", bootstyle="secondary")
lbl_pdf_template.grid(row=2, column=1, padx=10, sticky="w")

root.mainloop()