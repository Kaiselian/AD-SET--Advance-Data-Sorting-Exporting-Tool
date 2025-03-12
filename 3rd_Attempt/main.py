import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
import os
from file_reader import read_excel_csv
from docx_filler import fill_docx_template
from pdf_generator import generate_pdfs
from data_mapper import map_data_to_docx
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


# Initialize GUI
root = tb.Window(themename="journal")
root.title("Automated Document Filler")
root.geometry("1000x600")

# Global Variables
input_file = None
template_file = None
output_folder = None

def upload_data_file():
    global input_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.xls;*.csv")])
    if file_path:
        input_file = file_path
        lbl_data.config(text=f"📂 {os.path.basename(file_path)} Loaded")
        messagebox.showinfo("Success", "Data file loaded successfully!")

# 🟢 Load DOCX Template
def upload_template():
    global template_file
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        template_file = file_path
        lbl_template.config(text=f"📄 {os.path.basename(file_path)} Loaded")
        messagebox.showinfo("Success", "Template loaded successfully!")

# 🟢 Load PDF Template
def upload_pdf_template():
    global pdf_template_file
    file_path = filedialog.askopenfilename(filetypes=[("PDF Documents", "*.pdf")])
    if file_path:
        pdf_template_file = file_path
        lbl_pdf_template.config(text=f"📄 {os.path.basename(file_path)} Loaded")
        messagebox.showinfo("Success", "PDF Template loaded successfully!")

# 🟢 Select Output Folder
def select_output_folder():
    global output_folder
    folder = filedialog.askdirectory()
    if folder:
        output_folder = folder
        lbl_output.config(text=f"📁 Output Folder: {folder}")

def replace_pdf_placeholders(input_pdf, output_pdf, data):
    """Replaces placeholders in a PDF using overlay."""
    try:
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        page = reader.pages[0]

        writer.add_page(page)
        c = canvas.Canvas("temp_overlay.pdf", pagesize=letter)
        c.setFont("Helvetica", 12)

        # Placeholder positions (adjust per template)
        positions = {
            "{{Order:}}": (100, 700),
            "{{Order Date:}}": (300, 700),
            "{{Total: Quantity:}}": (500, 700),
        }

        for placeholder, value in data.items():
            if placeholder in positions:
                x, y = positions[placeholder]
                c.drawString(x, y, str(value))

        c.save()

        overlay_reader = PdfReader("temp_overlay.pdf")
        overlay_page = overlay_reader.pages[0]
        page.merge_page(overlay_page)

        with open(output_pdf, "wb") as f:
            writer.write(f)

        os.remove("temp_overlay.pdf")
        print(f"✅ Modified PDF saved: {output_pdf}")

    except Exception as e:
        print(f"❌ Error: {e}")

def draw_text_in_box(canvas, text, x, y, width, height, font_name="Helvetica", font_size=12):
    """Draws text within a specified box, handling overflow."""
    # ... (draw_text_in_box function from previous example) ...
    canvas.setFont(font_name, font_size)
    available_width = width
    lines = []
    words = text.split()
    current_line = ""

    for word in words:
        test_line = current_line + " " + word if current_line else word
        text_width = canvas.stringWidth(test_line, font_name, font_size)
        if text_width <= available_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word

    lines.append(current_line)

    line_height = font_size * 1.2  # Add some spacing
    for i, line in enumerate(lines):
        if (i + 1) * line_height <= height:
            canvas.drawString(x, y - i * line_height, line)
        else:
            canvas.drawString(x, y - i * line_height, line + "...") #add elipsis if truncated.
            break

# 🟢 Start Automated Processing
def start_processing():
    if not input_file or (not template_file and not pdf_template_file) or not output_folder:
        messagebox.showerror("Error", "Please upload all required files!")
        return

    # Progress Bar
    progress_window = tk.Toplevel(root)
    progress_window.title("Processing...")
    progress_bar = tb.Progressbar(progress_window, mode="indeterminate")
    progress_bar.pack(pady=20, padx=20)
    progress_bar.start()

    # Step 1: Read Data
    data = read_excel_csv(input_file)

    if data is None:
        messagebox.showerror("Error", "Failed to read data file.")
        progress_window.destroy()
        return

    if template_file: #docx processing
        # Step 2: Map Data to DOCX Placeholders
        mapped_data = map_data_to_docx(template_file, data)

        if mapped_data is None:
            messagebox.showerror("Error", "Failed to map data to template.")
            progress_window.destroy()
            return

        # Step 3: Fill DOCX Template
        filled_files = fill_docx_template(template_file, mapped_data, output_folder)

        if filled_files is None:
            messagebox.showerror("Error", "Failed to fill DOCX template.")
            progress_window.destroy()
            return

        # Step 4: Convert to PDFs
        generate_pdfs(filled_files, output_folder)

        progress_bar.stop()
        progress_window.destroy()

        messagebox.showinfo("Success", "All PDFs generated successfully!")
    elif pdf_template_file: #pdf processing
        # Step 2: Map Data to PDF Placeholders
        pdf_data = {}
        for column in data.columns:
            pdf_data[f"{{{{{column.strip()}}}}}"] = str(data[column].iloc[0])

        # Step 3: Replace placeholders in PDF
        output_pdf_file = os.path.join(output_folder, "output_invoice.pdf") #create output file name.
        replace_pdf_placeholders(pdf_template_file, output_pdf_file, pdf_data)

        progress_bar.stop()
        progress_window.destroy()

        messagebox.showinfo("Success", "PDF generated successfully!")

# 🔹 GUI Layout
frame = tb.Frame(root)
frame.pack(pady=20)

btn_data = tb.Button(frame, text="📂 Upload Data File", command=upload_data_file)
btn_data.grid(row=0, column=0, padx=10, pady=5)

btn_template = tb.Button(frame, text="📄 Upload DOCX Template", command=upload_template)
btn_template.grid(row=1, column=0, padx=10, pady=5)

btn_pdf_template = tb.Button(frame, text="📄 Upload PDF Template", command=upload_pdf_template) #add pdf template button.
btn_pdf_template.grid(row=2, column=0, padx=10, pady=5)

btn_output = tb.Button(frame, text="📁 Select Output Folder", command=select_output_folder)
btn_output.grid(row=3, column=0, padx=10, pady=5)

btn_start = tb.Button(frame, text="🚀 Start Processing", bootstyle="success", command=start_processing)
btn_start.grid(row=4, column=0, padx=10, pady=20)

# Labels for file paths
lbl_data = tb.Label(frame, text="No Data File Loaded", bootstyle="secondary")
lbl_data.grid(row=0, column=1, padx=10, sticky="w")

lbl_template = tb.Label(frame, text="No Template File Loaded", bootstyle="secondary")
lbl_template.grid(row=1, column=1, padx=10, sticky="w")

lbl_pdf_template = tb.Label(frame, text="No PDF Template File Loaded", bootstyle="secondary") #add pdf label.
lbl_pdf_template.grid(row=2, column=1, padx=10, sticky="w")

lbl_output = tb.Label(frame, text="No Output Folder Selected", bootstyle="secondary")
lbl_output.grid(row=3, column=1, padx=10, sticky="w")

root.mainloop()