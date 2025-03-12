import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
import darkdetect
import pandas as pd
from widgets.search_widgets import SearchWidgets
from widgets.export_widgets import ExportWidgets
from views.pdf_view import PDFView
from utils.data_utils import filter_data, display_data
from utils.file_utils import upload_file, export_filtered_data
from utils.pdf_generator import generate_pdfs
import os
from tkinter import filedialog, messagebox

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Data Search & Export Tool 1.13.6.1")
        self.root.geometry("1920x1080")
        self.root.state("zoomed")
        self.df = None
        self.filtered_df = None
        self.sort_orders = {}
        self.pdf_path = None
        self.setup_ui()

    def setup_ui(self):
        self.setup_top_frame()
        self.setup_treeview()
        self.setup_menu()

    def setup_top_frame(self):
        top_frame = tb.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X, padx=20)

        upload_btn = tb.Button(top_frame, text="📂 Upload File", bootstyle="primary", command=self.upload_data_file)
        upload_btn.pack(side=tk.LEFT, padx=10)

        self.search_widgets = SearchWidgets(top_frame, self.tree, self.perform_search)
        self.export_widgets = ExportWidgets(top_frame, self.filtered_df, self.export_data)

        btn_load_pdf = ttk.Button(top_frame, text="📂 Load PDF", command=self.load_pdf)
        btn_load_pdf.pack(side=tk.LEFT, padx=5)

        pdf_to_excel_btn = tb.Button(top_frame, text="📥 PDF to Excel", bootstyle="info", command=self.convert_pdf_to_excel)
        pdf_to_excel_btn.pack(side=tk.RIGHT, padx=10)

        btn_export_pdf = tb.Button(top_frame, text="📤 Export PDFs", bootstyle="success", command=self.export_filled_pdfs)
        btn_export_pdf.pack(side=tk.RIGHT, padx=10)

        upload_docx_btn = tk.Button(top_frame, text="Upload DOCX", command=self.upload_docx_and_convert)
        upload_docx_btn.pack(side=tk.LEFT, padx=10)

    def setup_treeview(self):
        frame2 = tb.Frame(self.root)
        frame2.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        tree_scroll_y = tk.Scrollbar(frame2, orient="vertical")
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        tree_scroll_x = tk.Scrollbar(frame2, orient="horizontal")
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(frame2, style="Custom.Treeview", yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

    def setup_menu(self):
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Upload File", command=self.upload_data_file)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.root.config(menu=menu_bar)

    def upload_data_file(self):
        self.df = upload_file()
        if self.df is not None:
            self.search_widgets.update_columns(self.df.columns)
            self.filtered_df = self.df.copy()
            display_data(self.tree, self.df, self.sort_orders)

    def perform_search(self, search_query, sub_query, main_column, sub_column, filter_type):
        if self.df is None:
            messagebox.showerror("Error", "Please upload a file first.")
            return

        self.filtered_df = filter_data(self.df, search_query, sub_query, main_column, sub_column, filter_type)
        display_data(self.tree, self.filtered_df, self.sort_orders)

    def export_data(self, format):
        if self.filtered_df is None:
            messagebox.showerror("Error", "No filtered data to export.")
            return
        export_filtered_data(self.filtered_df, format)

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            pdf_window = tk.Toplevel(self.root)
            pdf_window.title("PDF Preview")
            PDFView(pdf_window, self.pdf_path).pack(fill=tk.BOTH, expand=True)

    def convert_pdf_to_excel(self):
        # Your PDF to Excel conversion logic here
        pass

    def export_filled_pdfs(self):
        # Your PDF export logic here
        pass

    def upload_docx_and_convert(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("DOCX files", "*.docx")])
        if not file_paths:
            return
        output_folder = "output_pdfs"
        if not os.path.exists(output_folder):
            try:
                os.makedirs(output_folder)
            except OSError as e:
                messagebox.showerror("Error", f"Could not create output folder: {e}")
                return
        pdf_files = generate_pdfs(file_paths, output_folder)
        if pdf_files:
            messagebox.showinfo("Success", f"Successfully converted {len(pdf_files)} files to PDF.")
        else:
            messagebox.showerror("Error", "Failed to generate PDFs.")

if __name__ == "__main__":
    theme = "darkly" if darkdetect.isDark() else "journal"
    root = tb.Window(themename=theme)
    app = MainApp(root)
    root.mainloop()