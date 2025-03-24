from PyQt5.QtWidgets import QFileDialog, QMessageBox
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

def upload_file():
    file_path, _ = QFileDialog.getOpenFileName(None, "Open File", "", "Excel/CSV Files (*.xlsx *.xls *.csv)")
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False) if file_path.endswith(".csv") else pd.read_excel(file_path, sheet_name=0)
            if df.empty:
                QMessageBox.showerror("Error", "Loaded file is empty or could not be read.")
                return None
            QMessageBox.showinfo("Success", "File uploaded successfully!")
            return df
        except Exception as e:
            QMessageBox.showerror("Error", f"Failed to load file: {e}")
            return None
    return None

def export_filtered_data(df, format):
    save_path, _ = QFileDialog.getSaveFileName(None, "Save File", "", f"{format.upper()} Files (*.{format})")
    if save_path:
        try:
            if format == "xlsx":
                df.to_excel(save_path, index=False)
            elif format == "csv":
                df.to_csv(save_path, index=False)
            elif format == "pdf":
                save_df_as_pdf(df, save_path)
            QMessageBox.showinfo("Success", f"Filtered data saved as {format.upper()} successfully!")
        except Exception as e:
            QMessageBox.showerror("Error", f"Failed to save file: {e}")

def save_df_as_pdf(df, save_path):
    doc = SimpleDocTemplate(save_path, pagesize=A4)
    elements = []
    data = [df.columns.tolist()] + df.astype(str).values.tolist()
    table = Table(data)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ])
    table.setStyle(style)
    elements.append(table)
    doc.build(elements)