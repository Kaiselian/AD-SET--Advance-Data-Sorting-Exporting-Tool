from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLineEdit, QMenu, QMenuBar, QAction, QFileDialog, QMessageBox, QTreeWidget, QTreeWidgetItem
from PyQt5.QtCore import Qt
import pandas as pd
import pdfplumber
import fitz
import os
from utils.file_utils import export_filtered_data
from utils.data_utils import filter_data
from utils.docx_filler import fill_docx_template
from views.pdf_view import PDFView
from utils.gui_utils import display_data, create_table_widget

class MainView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.df = None
        self.filtered_df = None
        self.sort_orders = {}  # Track sorting order for columns
        self.template_file = None
        self.output_folder = None
        self.pdf_path = None
        self.pdf_document = None
        self.box_data = None
        self.pdf_canvas = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Top Frame Layout
        top_frame = QWidget()
        top_layout = QHBoxLayout()

        # Upload Button
        self.upload_btn = QPushButton("📂 Upload File")
        self.upload_btn.clicked.connect(self.upload_file)
        top_layout.addWidget(self.upload_btn)

        # Search Entry
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Search...")
        self.search_entry.returnPressed.connect(self.search_and_generate)
        top_layout.addWidget(self.search_entry)

        # Search Button
        self.search_btn = QPushButton("🔍")
        self.search_btn.clicked.connect(self.search_and_generate)
        top_layout.addWidget(self.search_btn)

        # Sub-Search Entry
        self.sub_search_entry = QLineEdit()
        self.sub_search_entry.setPlaceholderText("Sub-Search...")
        self.sub_search_entry.returnPressed.connect(self.search_and_generate)
        top_layout.addWidget(self.sub_search_entry)

        # Sub-Search Column Dropdown
        self.sub_search_column_dropdown = QComboBox()
        self.sub_search_column_dropdown.addItem("All Columns")
        top_layout.addWidget(self.sub_search_column_dropdown)

        # Sub-Search Button
        self.sub_search_btn = QPushButton("🔍 Sub-Search")
        self.sub_search_btn.clicked.connect(self.search_and_generate)
        top_layout.addWidget(self.sub_search_btn)

        # Column Dropdown
        self.column_dropdown = QComboBox()
        self.column_dropdown.addItem("All Columns")
        top_layout.addWidget(self.column_dropdown)

        # Filter Type Dropdown
        self.filter_dropdown = QComboBox()
        self.filter_dropdown.addItems(["Contains", "Equals", "Starts with"])
        top_layout.addWidget(self.filter_dropdown)

        # Clear Button
        self.clear_btn = QPushButton("❌ Clear Filters")
        self.clear_btn.clicked.connect(self.clear_filters)
        top_layout.addWidget(self.clear_btn)

        # Load PDF Button
        self.load_pdf_btn = QPushButton("📂 Load PDF")
        self.load_pdf_btn.clicked.connect(self.load_pdf)
        top_layout.addWidget(self.load_pdf_btn)

        # PDF to Excel Button
        self.pdf_to_excel_btn = QPushButton("📥 PDF to Excel")
        self.pdf_to_excel_btn.clicked.connect(self.convert_pdf_to_excel)
        top_layout.addWidget(self.pdf_to_excel_btn)

        # Export Menu Button
        self.export_menu_btn = QPushButton("📤 Export")
        self.export_menu = QMenu(self)
        self.export_menu.addAction("📤 Export as CSV", lambda: self.export_filtered_data("csv"))
        self.export_menu.addAction("📤 Export as Excel", lambda: self.export_filtered_data("xlsx"))
        self.export_menu.addAction("📤 Export Full PDF", lambda: self.export_filtered_data("pdf"))
        self.export_menu.addAction("📤 Export Individual PDFs", self.export_each_row_as_pdf)
        self.export_menu_btn.setMenu(self.export_menu)
        top_layout.addWidget(self.export_menu_btn)

        # Export PDFs Button
        self.export_pdfs_btn = QPushButton("📤 Export PDFs")
        self.export_pdfs_btn.clicked.connect(self.export_filled_pdfs)
        top_layout.addWidget(self.export_pdfs_btn)

        top_frame.setLayout(top_layout)
        layout.addWidget(top_frame)

        # Treeview
        self.tree = create_table_widget(self)
        layout.addWidget(self.tree)

        self.setLayout(layout)

    def clear_filters(self):
        """Resets all search filters and refreshes the dataset."""
        if self.df is None:
            QMessageBox.showerror(self, "Error", "No data loaded to clear filters.")
            return

        self.filtered_df = self.df.copy()  # Reset data
        self.sort_orders = {}  # Reset sorting order

        self.search_entry.setText("")
        self.sub_search_entry.setText("")
        self.column_dropdown.setCurrentText("All Columns")
        self.sub_search_column_dropdown.setCurrentText("All Columns")
        self.filter_dropdown.setCurrentText("Contains")

        self.display_data(self.search_entry.text(), self.sub_search_entry.text(), self.column_dropdown.currentText(),
                          self.sub_search_column_dropdown.currentText(), self.filter_dropdown.currentText())

    def upload_file(self):
        """Handles the file upload."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if file_path:
            try:
                if file_path.endswith(".xlsx"):
                    self.df = pd.read_excel(file_path)
                elif file_path.endswith(".csv"):
                    self.df = pd.read_csv(file_path)
                self.column_dropdown.clear()
                self.column_dropdown.addItem("All Columns")
                self.column_dropdown.addItems(list(self.df.columns))
                self.sub_search_column_dropdown.clear()
                self.sub_search_column_dropdown.addItem("All Columns")
                self.sub_search_column_dropdown.addItems(list(self.df.columns))
                self.display_data()
            except Exception as e:
                QMessageBox.showerror(self, "Error", f"Failed to upload file: {e}")
        else:
            print("No file selected.")

    def display_data(self, search_query="", sub_query="", main_column="All Columns", sub_column="All Columns", filter_type="Contains"):
        """Filters and updates the Treeview based on search criteria."""
        if self.df is None:
            return

        # Apply filtering
        filtered_df = filter_data(self.df, search_query, sub_query, main_column, sub_column, filter_type)

        # Update Treeview
        display_data(self.tree, filtered_df, self.sort_orders)

    def open_pdf_view(self):
        """Opens the PDF view with the loaded PDF."""
        if self.pdf_path:
            pdf_window = QWidget()
            pdf_window.setWindowTitle("PDF Preview")
            pdf_view_instance = PDFView(pdf_window, self.pdf_path)
            layout = QVBoxLayout()
            layout.addWidget(pdf_view_instance)
            pdf_window.setLayout(layout)
            pdf_window.show()
        else:
            QMessageBox.showerror(self, "Error", "No PDF loaded.")

    def start_processing(self):
        """Handles the processing of data and template files."""
        if not self.df or not self.template_file or not self.output_folder:
            QMessageBox.showerror(self, "Error", "Please upload all required files!")
            return

        # Fill DOCX Templates for Each Row
        filled_files =