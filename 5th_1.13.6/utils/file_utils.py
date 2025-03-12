import pandas as pd
import os
from tkinter import filedialog, messagebox

def upload_file():
    """Uploads an Excel or CSV file and returns a DataFrame."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False) if file_path.endswith(".csv") else pd.read_excel(file_path, sheet_name=0)

            if df.empty:
                messagebox.showerror("Error", "Loaded file is empty or could not be read.")
                return None

            messagebox.showinfo("Success", "File uploaded successfully!")
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            return None
    return None

def export_filtered_data(df, format):
    """Exports filtered data to the specified format (CSV, Excel, PDF)."""
    if df is None or df.empty:
        messagebox.showerror("Error", "No data to export.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=f".{format}", filetypes=[(f"{format.upper()} files", f"*.{format}")])
    if save_path:
        try:
            if format == "xlsx":
                df.to_excel(save_path, index=False)
            elif format == "csv":
                df.to_csv(save_path, index=False)
            elif format == "pdf":
                save_df_as_pdf(df, save_path)
            messagebox.showinfo("Success", f"Filtered data saved as {format.upper()} successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

def save_df_as_pdf(df, save_path):
    """Saves a pandas DataFrame as a PDF."""
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        # Add DataFrame header
        for col in df.columns:
            pdf.cell(40, 10, txt=str(col), border=1)
        pdf.ln()

        # Add DataFrame rows
        for _, row in df.iterrows():
            for item in row:
                pdf.cell(40, 10, txt=str(item), border=1)
            pdf.ln()

        pdf.output(save_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save PDF: {e}")
        return False
    return True