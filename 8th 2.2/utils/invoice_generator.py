import os
import pandas as pd
from fpdf import FPDF
from PyQt5.QtWidgets import QFileDialog, QMessageBox


class InvoiceGenerator:
    def __init__(self, template_path=None):
        """Initialize invoice generator with an optional PDF template."""
        self.template_path = template_path
        self.df = None  # Store the loaded data

    def load_invoice_data(self, file_path=None):
        """Loads invoice data from an Excel/CSV file."""
        if file_path is None:
            file_path, _ = QFileDialog.getOpenFileName(None, "Open File", "", "Excel/CSV Files (*.xlsx *.xls *.csv)") # Change 1

        if not file_path:
            QMessageBox.showerror("Error", "No file selected!") # Change 2
            return None

        try:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)

            # Debug: Print the DataFrame
            print("DataFrame loaded:")
            print(self.df)

            if self.df.empty:
                QMessageBox.showerror("Error", "The file is empty!") # Change 2
                return None

            QMessageBox.showinfo("Success", f"Loaded {len(self.df)} invoice records!") # Change 2
            return self.df

        except Exception as e:
            QMessageBox.showerror("Error", f"Failed to load file: {e}") # Change 2
            return None

    def generate_invoices(self, output_folder=None):
        """Generates PDF invoices from the dataset."""
        if self.df is None or self.df.empty:
            QMessageBox.showerror("Error", "No data loaded for invoices!") # Change 2
            return

        if output_folder is None:
            output_folder = QFileDialog.getExistingDirectory(None, "Select Output Folder") # Change 1

        if not output_folder:
            QMessageBox.showerror("Error", "No output folder selected!") # Change 2
            return

        # Debug: Print the output folder
        print(f"Output folder: {output_folder}")

        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Generate a PDF for each row of data (starting from the second row)
        for index, row in self.df.iloc[1:].iterrows():
            self.create_invoice(row, output_folder, index)

        QMessageBox.showinfo("Success", f"Invoices saved in {output_folder}!") # Change 2

    def create_invoice(self, data_row, output_folder, invoice_number):
        """Creates a single invoice PDF."""
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        # Add invoice title
        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, "Invoice", ln=True, align='C')

        # Add client details
        pdf.set_font("Arial", size=12)
        pdf.cell(100, 10, f"Client: {data_row.get(self.df.columns[0], 'Unknown')}", ln=True)
        pdf.cell(100, 10, f"Invoice #: {invoice_number}", ln=True)

        # Add a space
        pdf.cell(200, 10, "", ln=True)

        # Add table headers
        pdf.set_font("Arial", "B", 12)
        for col in self.df.columns:
            pdf.cell(40, 10, col, 1)
        pdf.ln()

        # Add table data
        pdf.set_font("Arial", size=12)
        for col in self.df.columns:
            pdf.cell(40, 10, str(data_row[col]), 1)
        pdf.ln()

        # Save the PDF
        output_path = os.path.join(output_folder, f"Invoice_{invoice_number}.pdf")
        pdf.output(output_path)

        # Debug: Verify the PDF was created
        if os.path.exists(output_path):
            print(f"✅ Invoice saved: {output_path}")
        else:
            print(f"❌ Failed to save invoice: {output_path}")