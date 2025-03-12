import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
from utils.pdf_utils import load_pdf

class PDFView(tb.Toplevel):  # Inherit from tb.Toplevel
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.parent = parent
        self.pdf_path = pdf_path
        self.title("PDF Preview")  # Set the window title
        self.geometry("1200x900")
        self.create_widgets()

    def create_widgets(self):
        """Creates the widgets for the PDF preview window."""
        # Load PDF
        self.pdf_img, self.pdf_document = load_pdf(self.pdf_path)

        # Canvas for PDF Preview
        self.canvas = tk.Canvas(self, width=self.pdf_img.width(), height=self.pdf_img.height())
        self.canvas.pack(side=tk.LEFT, expand=True)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_img)

        # Frame for Right-Side Controls
        self.frame_right = tb.Frame(self, width=300, bg="#f0f0f0")
        self.frame_right.pack(side=tk.RIGHT, fill=tk.Y)

        # Add Text Box Button
        self.btn_add_box = tb.Button(self.frame_right, text="➕ Add Text Box", command=self.add_text_box)
        self.btn_add_box.pack(pady=10, padx=10, fill=tk.X)

    def add_text_box(self):
        """Adds a text box to the PDF preview."""
        frame = tk.Frame(self)
        entry = tk.Entry(frame, font=("Arial", 12), width=15)
        entry.pack(side=tk.LEFT)

        box_window = self.canvas.create_window(50, 50 + len(self.text_boxes) * 30, window=frame, anchor=tk.NW)
        self.text_boxes.append(entry)
        self.box_data.append({"entry": entry, "window": box_window, "x": 50, "y": 50, "column": None})