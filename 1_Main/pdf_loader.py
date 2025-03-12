# pdf_loader.py
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk


def open_pdf_file():
    """Open a file dialog to select a PDF and return the file path."""
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return pdf_path


def load_pdf_preview(pdf_path):
    """Load the first page of the PDF and return an ImageTk object."""
    if not pdf_path:
        return None, None, None

    pdf_document = fitz.open(pdf_path)  # Load PDF into memory
    page = pdf_document[0]  # First page
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.thumbnail((800, 1000))  # Resize for display
    pdf_img = ImageTk.PhotoImage(img)

    return pdf_document, pdf_img, img.width, img.height


def create_pdf_window(root, pdf_img, img_width, img_height, export_filled_pdfs, extract_to_excel, set_extraction_start,
                      update_columns):
    """Create a new window to display the PDF preview and controls."""
    pdf_window = tk.Toplevel(root)
    pdf_window.title("PDF Preview - Assign Data Fields")
    pdf_window.geometry("1200x900")
    pdf_window.state("zoomed")

    # 🟢 Right Panel (Buttons & Controls)
    frame_right = tk.Frame(pdf_window, width=300, bg="#f0f0f0")
    frame_right.pack(side=tk.RIGHT, fill=tk.Y)

    btn_export = tk.Button(frame_right, text="📤 Export PDFs", command=export_filled_pdfs)
    btn_export.pack(pady=10, padx=10, fill=tk.X)

    btn_extract = tk.Button(frame_right, text="📄 Extract & Export to Excel", command=extract_to_excel)
    btn_extract.pack(pady=10, padx=10, fill=tk.X)

    btn_set_segment = tk.Button(frame_right, text="📍 Set Extraction Start", command=set_extraction_start)
    btn_set_segment.pack(pady=10, padx=10, fill=tk.X)

    column_var = tk.StringVar()
    column_dropdown = ttk.Combobox(frame_right, textvariable=column_var, state="readonly")
    column_dropdown.pack(pady=5, padx=10, fill=tk.X)
    update_columns()

    # 🟢 Left Panel (PDF Canvas)
    pdf_canvas = tk.Canvas(pdf_window, width=img_width, height=img_height)
    pdf_canvas.pack(side=tk.LEFT, expand=True)
    pdf_canvas.create_image(0, 0, anchor=tk.NW, image=pdf_img)

    return pdf_canvas


def load_pdf(root, export_filled_pdfs, extract_to_excel, set_extraction_start, update_columns):
    """Main function to handle PDF loading and preview."""
    pdf_path = open_pdf_file()
    if not pdf_path:
        return

    pdf_document, pdf_img, img_width, img_height = load_pdf_preview(pdf_path)
    if pdf_document and pdf_img:
        pdf_canvas = create_pdf_window(root, pdf_img, img_width, img_height, export_filled_pdfs, extract_to_excel,
                                       set_extraction_start, update_columns)
        return pdf_path, pdf_document, pdf_canvas
