# main.py
import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QLabel, QTableWidget, QTableWidgetItem, QLineEdit,
    QComboBox, QDialog, QListWidget, QListWidgetItem, QFormLayout, QDialogButtonBox,
    QScrollArea, QGraphicsView, QGraphicsScene, QGraphicsRectItem
)
from PyQt5.QtCore import Qt, QFileInfo, QStandardPaths
from PyQt5.QtGui import QPixmap
from fpdf import FPDF
import fitz
import json
import logging
from utils.data_mapper import DataMapper
from docx import Document

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced Data Search & Export Tool 2.2")
        self.setGeometry(100, 100, 1200, 800)
        self.init_data()
        self.init_ui()

    def init_data(self):
        self.df = None
        self.filtered_df = None
        self.pdf_path = None
        self.pdf_document = None
        self.docx_template_path = None
        self.image_path = None
        self.box_column_map = {}

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.create_top_bar()
        self.create_search_widgets()
        self.create_data_table()
        self.create_docx_controls()
        self.create_pdf_preview()

    def create_top_bar(self):
        self.top_bar_layout = QHBoxLayout()
        self.layout.addLayout(self.top_bar_layout)

        buttons = [
            ("Load Data", self.load_data),
            ("Generate All Invoices", self.generate_all_invoices),
            ("Create Invoice", self.create_invoice_dialog),
            ("Load Image", self.load_image),
            ("Add Box", self.add_box),
            ("Save Structure", self.save_structure),
            ("Load Structure", self.load_structure),
            ("Export as CSV", lambda: self.export_data("csv")),
            ("Export as Excel", lambda: self.export_data("xlsx")),
            ("Export as PDF", lambda: self.export_data("pdf")),
            ("📂 Upload DOCX Template", self.upload_template)
        ]

        for text, callback in buttons:
            btn = QPushButton(text)
            btn.clicked.connect(callback)
            self.top_bar_layout.addWidget(btn)

    def create_search_widgets(self):
        self.label = QLabel("Load an Excel/CSV file to search, filter, and export data.")
        self.layout.addWidget(self.label)

        self.search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_layout.addWidget(self.search_input)

        self.filter_column = QComboBox()
        self.filter_column.addItem("All Columns")
        self.search_layout.addWidget(self.filter_column)

        self.filter_type = QComboBox()
        self.filter_type.addItems(["Contains", "Equals", "Starts with"])
        self.search_layout.addWidget(self.filter_type)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.perform_search)
        self.search_layout.addWidget(self.search_button)

        self.layout.addLayout(self.search_layout)

    def create_data_table(self):
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

    def create_docx_controls(self):
        self.fill_docx_btn = QPushButton("📝 Fill DOCX Template")
        self.fill_docx_btn.clicked.connect(self.fill_docx_template)
        self.layout.addWidget(self.fill_docx_btn)

    def create_pdf_preview(self):
        self.pdf_container = QWidget()
        self.pdf_layout = QHBoxLayout(self.pdf_container)

        self.graphics_scene = QGraphicsScene()
        self.graphics_view = QGraphicsView(self.graphics_scene)
        self.pdf_layout.addWidget(self.graphics_view)

        self.right_panel = QWidget()
        self.right_layout = QVBoxLayout(self.right_panel)
        self.pdf_layout.addWidget(self.right_panel)

        self.layout.addWidget(self.pdf_container)

    def load_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "Excel/CSV Files (*.xlsx *.xls *.csv)"
        )

        if not file_path:
            QMessageBox.warning(self, "Error", "No file selected!")
            return

        try:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)

            if self.df.empty:
                QMessageBox.warning(self, "Error", "The file is empty!")
                return

            self.filter_column.clear()
            self.filter_column.addItem("All Columns")
            self.filter_column.addItems(self.df.columns.tolist())

            self.display_data_in_table(self.df)
            QMessageBox.information(self, "Success", f"Loaded {len(self.df)} records!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file: {e}")

    def display_data_in_table(self, data):
        self.table.setRowCount(data.shape[0])
        self.table.setColumnCount(data.shape[1])
        self.table.setHorizontalHeaderLabels(data.columns)

        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                self.table.setItem(i, j, QTableWidgetItem(str(data.iat[i, j])))

    def perform_search(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data loaded!")
            return

        search_query = self.search_input.text().strip()
        filter_column = self.filter_column.currentText()
        filter_type = self.filter_type.currentText()

        if not search_query:
            self.display_data_in_table(self.df)
            return

        sub_queries = [q.strip() for q in search_query.split(',')]
        filtered_data = self.df.copy()

        for q in sub_queries:
            if filter_column == "All Columns":
                filtered_data = filtered_data[
                    filtered_data.apply(lambda row: row.astype(str).str.contains(q, case=False, na=False).any(), axis=1)
                ]
            else:
                if filter_type == "Contains":
                    filtered_data = filtered_data[filtered_data[filter_column].astype(str).str.contains(q, case=False, na=False)]
                elif filter_type == "Equals":
                    filtered_data = filtered_data[filtered_data[filter_column].astype(str) == q]
                elif filter_type == "Starts with":
                    filtered_data = filtered_data[filtered_data[filter_column].astype(str).str.startswith(q, na=False)]

        self.filtered_df = filtered_data

        if self.filtered_df.empty:
            QMessageBox.information(self, "No Results", "No matching records found.")
            return

        self.display_data_in_table(self.filtered_df)

    def export_data(self, format):
        if self.filtered_df is None or self.filtered_df.empty:
            QMessageBox.warning(self, "Error", "No filtered data to export!")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save File", "", f"{format.upper()} Files (*.{format})"
        )

        if not file_path:
            return

        try:
            if format == "csv":
                self.filtered_df.to_csv(file_path, index=False)
            elif format == "xlsx":
                self.filtered_df.to_excel(file_path, index=False)
            elif format == "pdf":
                self.save_df_as_pdf(self.filtered_df, file_path)

            QMessageBox.information(self, "Success", f"Data exported as {format.upper()} successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export data: {e}")

    def save_df_as_pdf(self, df, save_path):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, "Exported Data", ln=True, align='C')

        pdf.set_font("Arial", "B", 12)
        for col in df.columns:
            pdf.cell(40, 10, col, 1)
        pdf.ln()

        pdf.set_font("Arial", size=12)
        for _, row in df.iterrows():
            for col in df.columns:
                pdf.cell(40, 10, str(row[col]), 1)
            pdf.ln()

        pdf.output(save_path)
        logger.info(f"PDF saved: {save_path}")

    def create_invoice_dialog(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data loaded!")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Create Invoice")
        layout = QVBoxLayout(dialog)

        columns_label = QLabel("Select Columns:")
        layout.addWidget(columns_label)
        columns_list = QListWidget()
        columns_list.setSelectionMode(QListWidget.MultiSelection)
        for col in self.df.columns:
            item = QListWidgetItem(col)
            columns_list.addItem(item)
        layout.addWidget(columns_list)

        details_label = QLabel("Invoice Details:")
        layout.addWidget(details_label)
        form_layout = QFormLayout()
        invoice_number = QLineEdit()
        form_layout.addRow("Invoice Number:", invoice_number)
        customer_name = QLineEdit()
        form_layout.addRow("Customer Name:", customer_name)
        invoice_date = QLineEdit()
        form_layout.addRow("Invoice Date:", invoice_date)
        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        dialog.exec_()

    def generate_all_invoices(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data loaded!")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not output_folder:
            return

        for index, row in self.df.iterrows():
            invoice_number = str(row.get("Invoice Number", f"INV-{index + 1}"))
            customer_name = str(row.get("Customer Name", f"Customer {index + 1}"))
            invoice_date = str(row.get("Invoice Date", "N/A"))

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Invoice", ln=True, align='C')
            pdf.set_font("Arial", size=12)

            address_columns = ["Address", "City", "Zip"]
            if address_columns:
                pdf.cell(0, 10, "Customer Information:", ln=True)
                address_text = ""
                for col in address_columns:
                    if col in row:
                        address_text += f"{col}: {row[col]}\n"
                pdf.multi_cell(0, 10, address_text)

            pdf.cell(0, 10, f"Invoice Number: {invoice_number}", ln=True)
            pdf.cell(0, 10, f"Customer Name: {customer_name}", ln=True)
            pdf.cell(0, 10, f"Invoice Date: {invoice_date}", ln=True)
            pdf.ln(10)

            pdf.set_font("Arial", "B", 12)
            for col in self.df.columns:
                pdf.cell(40, 10, col, 1)
            pdf.ln()

            pdf.set_font("Arial", size=12)
            for col in self.df.columns:
                pdf.cell(40, 10, str(row[col]), 1)
            pdf.ln()

            file_path = os.path.join(output_folder, f"Invoice_{index + 1}.pdf")
            pdf.output(file_path)

        QMessageBox.information(self, "Success", f"{len(self.df)} invoices generated successfully!")

    def load_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Image", "", "Image Files (*.png *.jpg *.jpeg)")
        if file_path:
            self.image_path = file_path
            self.create_pdf_preview()

    def create_pdf_preview(self):
        if self.image_path:
            pixmap = QPixmap(self.image_path)
            self.pdf_label = QLabel()
            self.pdf_label.setPixmap(pixmap)
            self.scroll_area = QScrollArea()
            self.scroll_area.setWidget(self.pdf_label)
            self.graphics_scene = QGraphicsScene()
            self.graphics_view = QGraphicsView(self.graphics_scene)
            self.graphics_view.setGeometry(self.scroll_area.geometry())
            self.graphics_view.setStyleSheet("background: transparent;")
            self.graphics_view.setAttribute(Qt.WA_TranslucentBackground)
            self.layout.addWidget(self.scroll_area)
            self.layout.addWidget(self.graphics_view)
            for rect, dropdown in self.box_column_map.items():
                self.graphics_scene.addItem(rect)
                self.layout.addWidget(dropdown)

    def add_box(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data loaded!")
            return

        rect = QGraphicsRectItem(100, 100, 200, 50)
        self.graphics_scene.addItem(rect)
        column_dropdown = QComboBox()
        column_dropdown.addItems(self.df.columns.tolist())
        self.layout.addWidget(column_dropdown)
        self.box_column_map[rect] = column_dropdown

    def save_structure(self):
        structure = []
        for rect, dropdown in self.box_column_map.items():
            structure.append({
                "x": rect.rect().x(),
                "y": rect.rect().y(),
                "width": rect.rect().width(),
                "height": rect.rect().height(),
                "column": dropdown.currentText()
            })
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Structure", "", "JSON Files (*.json)")
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(structure, f)

    def load_structure(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Load Structure", "", "JSON Files (*.json)")
        if file_path:
            with open(file_path, 'r') as f:
                structure = json.load(f)
            self.box_column_map = {}
            self.graphics_scene.clear()
            for item in structure:
                rect = QGraphicsRectItem(item['x'], item['y'], item['width'], item['height'])
                self.graphics_scene.addItem(rect)
                column_dropdown = QComboBox()
                column_dropdown.addItems(self.df.columns.tolist())
                column_dropdown.setCurrentText(item['column'])
                self.layout.addWidget(column_dropdown)
                self.box_column_map[rect] = column_dropdown
            self.create_pdf_preview()

    def generate_pdf_with_boxes(self, output_path):
        doc = fitz.open()
        page = doc.new_page()
        if self.image_path:
            rect = page.rect
            page.insert_image(rect, filename=self.image_path)
        for rect, column_dropdown in self.box_column_map.items():
            column_name = column_dropdown.currentText()
            text = str(self.df.iloc[0][column_name])
            x = rect.rect().x()
            y = rect.rect().y()
            page.insert_text((x, y), text)
        doc.save(output_path)

    def upload_template(self):
        docs_path = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)

        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select DOCX Template",
                docs_path,
                "Word Documents (*.docx);;All Files (*)"
            )

            if not file_path:
                return False

            file_info = QFileInfo(file_path)

            if not file_info.exists():
                QMessageBox.critical(self, "File Not Found", "The selected file does not exist.")
                return False

            max_size_mb = 20
            file_size_mb = file_info.size() / (1024 * 1024)
            if file_size_mb > max_size_mb:
                QMessageBox.critical(self, "File Too Large", f"File exceeds maximum size of {max_size_mb}MB")
                return False

            if not file_info.isReadable():
                QMessageBox.critical(self, "Permission Denied", "You don't have permission to read this file.")
                return False

            try:
                doc = Document(file_path)
                if not doc.paragraphs and not doc.tables:
                    QMessageBox.warning(self, "Empty Document", "The document appears to be empty or corrupted.")
                    return False

                temp_path = os.path.join(QStandardPaths.writableLocation(
                    QStandardPaths.TempLocation), "temp_validation.docx")
                doc.save(temp_path)
                os.remove(temp_path)

                self.docx_template_path = file_path
                QMessageBox.information(self, "Success", "DOCX template loaded successfully!")
                return True

            except Exception as e:
                QMessageBox.critical(self, "Document Error", f"Failed to process document: {str(e)}")
                return False

        except Exception as e:
            QMessageBox.critical(self, "Unexpected Error", f"An unexpected error occurred:\n{str(e)}")
            return False

    def fill_docx_template(self):
        if not hasattr(self, 'docx_template_path') or not self.docx_template_path:
            QMessageBox.critical(self, "Error", "No DOCX template uploaded!")
            return

        if self.df is None or self.df.empty:
            QMessageBox.critical(self, "Error", "No data loaded!")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not output_folder:
            return

        try:
            mapper = DataMapper()  # Remove self
            result = mapper.map_data_to_docx(
                self.docx_template_path,
                self.df,
                output_folder
            )

            if result:
                QMessageBox.information(
                    self,
                    "Success",
                    f"Successfully generated {len(result)} documents in:\n{output_folder}"
                )
            else:
                QMessageBox.warning(
                    self,
                    "Warning",
                    "Documents were not generated. Please check the logs."
                )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to generate documents:\n{str(e)}"
            )
            logger.error(f"Document generation failed: {str(e)}", exc_info=True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())