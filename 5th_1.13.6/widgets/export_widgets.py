import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
from utils.file_utils import export_filtered_data
from utils.pdf_generator import generate_pdfs

class ExportWidgets:
    def __init__(self, parent, df):
        self.parent = parent
        self.df = df
        self.create_widgets()

    def create_widgets(self):
        # Export Menu Button
        self.export_menu_btn = tb.Menubutton(self.parent, text="📤 Export", bootstyle="warning")
        self.export_menu_btn.pack(side=tk.RIGHT, padx=10)

        # Create the Dropdown Menu
        self.export_menu = tk.Menu(self.export_menu_btn, tearoff=0)
        self.export_menu.add_command(label="📤 Export as CSV", command=lambda: export_filtered_data(self.df, "csv"))
        self.export_menu.add_command(label="📤 Export as Excel", command=lambda: export_filtered_data(self.df, "xlsx"))
        self.export_menu.add_command(label="📤 Export as PDF", command=lambda: export_filtered_data(self.df, "pdf"))

        # Attach the Menu to the Button
        self.export_menu_btn["menu"] = self.export_menu

docx_files = ["file1.docx", "file2.docx"]
output_folder = "output_pdfs"
pdf_files = generate_pdfs(docx_files, output_folder)

if pdf_files:
    print(f"Generated PDFs: {pdf_files}")
else:
    print("Failed to generate PDFs.")