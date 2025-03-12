# file_loader.py
import pandas as pd
from tkinter import filedialog, messagebox

def select_file():
    """Open a file dialog and return the selected file path."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    return file_path

def load_dataframe(file_path):
    """Load data from CSV or Excel into a pandas DataFrame."""
    if not file_path:
        return None, "No file selected."

    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False)
        else:
            df = pd.read_excel(file_path, sheet_name=0)

        if df.empty:
            return None, "Loaded file is empty or could not be read."

        return df, None  # Return dataframe and no error
    except Exception as e:
        return None, f"Failed to load file: {e}"

def upload_file(update_columns, display_data):
    """Handle file upload, update UI, and display data."""
    file_path = select_file()
    df, error = load_dataframe(file_path)

    if error:
        messagebox.showerror("Error", error)
        print("Upload Error:", error)  # Debugging
        return None

    update_columns()
    display_data(df)
    messagebox.showinfo("Success", "File uploaded successfully!")

    return df
