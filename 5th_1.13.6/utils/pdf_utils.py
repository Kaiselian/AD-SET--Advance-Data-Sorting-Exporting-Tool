import fitz  # PyMuPDF
from PIL import Image, ImageTk
import ttkbootstrap as tb
from tkinter import messagebox
import tkinter as tk
import re
import os

def load_pdf(pdf_path):
    """Loads a PDF and returns a tkinter PhotoImage and the PDF document."""
    if not pdf_path: #add this check.
        messagebox.showerror("Error", "No PDF path provided.")
        return None, None
    try:
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[0]  # Get the first page
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        pdf_img = ImageTk.PhotoImage(img)
        return pdf_img, pdf_document
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load PDF: {e}")
        return None, None

def extract_placeholders_from_pdf(pdf_path):
    """Extracts placeholders from a PDF."""
    placeholders = set()
    if not pdf_path: #add this check.
        messagebox.showerror("Error", "No PDF path provided.")
        return []
    try:
        pdf_document = fitz.open(pdf_path)
        for page in pdf_document:
            text = page.get_text()
            matches = re.findall(r"\{\{.*?\}\}", text)
            for match in matches:
                placeholders.add(match)
        pdf_document.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract placeholders: {e}")
    return list(placeholders)

class MainView(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.pdf_path = None
        self.create_widgets()

    def create_widgets(self):

        # Load PDF Button
        self.load_pdf_btn = tb.Button(self, text="Load PDF", command=self.load_pdf_file)
        self.load_pdf_btn.pack(pady=10)

    def open_pdf_view(self):
        """Opens the PDF view with the loaded PDF."""
        if self.pdf_path:
            pdf_window = tk.Toplevel(self.parent)
            pdf_window.title("PDF Preview")
            PDFView(pdf_window, self.pdf_path).pack(fill=tk.BOTH, expand=True)  # pass in the pdf path.
        else:
            messagebox.showerror("Error", "No PDF loaded.")

    class MainView(tb.Frame):  # This is likely a copy paste error, and should be removed.
        def __init__(self, parent):
            super().__init__(parent)

# Example PDFView class (adjust to your actual implementation)
class PDFView(tb.Frame):
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.create_widgets()

    def create_widgets(self):
        self.pdf_img, self.pdf_document = load_pdf(self.pdf_path)
        if self.pdf_img is None:
            return

        self.canvas = tk.Canvas(self)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbars
        v_scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar = tk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.canvas.config(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.canvas.config(scrollregion=(0, 0, self.pdf_img.width(), self.pdf_img.height()))

        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_img)
