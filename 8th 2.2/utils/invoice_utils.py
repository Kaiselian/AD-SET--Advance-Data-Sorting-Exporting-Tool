from fpdf import FPDF
from PyQt5.QtWidgets import QMessageBox, QFileDialog

def generate_pdf_invoice(df, selected_columns, invoice_number, customer_name, invoice_date, parent):
    """Generates a PDF invoice from a DataFrame."""
    if not selected_columns:
        QMessageBox.warning(parent, "Error", "Please select at least one column.")
        return False

    selected_columns_names = [item.text() for item in selected_columns]
    invoice_df = df[selected_columns_names].copy()

    # Create PDF Invoice
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, "Invoice", ln=True, align='C')

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, f"Invoice Number: {invoice_number}", ln=True)
    pdf.cell(200, 10, f"Customer Name: {customer_name}", ln=True)
    pdf.cell(200, 10, f"Invoice Date: {invoice_date}", ln=True)
    pdf.ln(10)

    # Add table headers
    pdf.set_font("Arial", "B", 12)
    for col in invoice_df.columns:
        pdf.cell(40, 10, col, 1)
    pdf.ln()

    # Add table data
    pdf.set_font("Arial", size=12)
    for _, row in invoice_df.iterrows():
        for col in invoice_df.columns:
            pdf.cell(40, 10, str(row[col]), 1)
        pdf.ln()

    # Save the PDF
    file_path, _ = QFileDialog.getSaveFileName(parent, "Save Invoice", "", "PDF Files (*.pdf)")
    if file_path:
        pdf.output(file_path)
        QMessageBox.information(parent, "Success", "Invoice generated successfully!")
        return True
    else:
        QMessageBox.information(parent, "Cancelled", "Invoice generation cancelled.")
        return False