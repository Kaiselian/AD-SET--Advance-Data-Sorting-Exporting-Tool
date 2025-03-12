# data_loader.py
import pandas as pd
from tkinter import filedialog, messagebox


def upload_file(update_columns, display_data):
    """Uploads an Excel or CSV file and loads it into a DataFrame."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])

    if not file_path:
        return None  # User canceled file selection

    try:
        df = pd.read_csv(file_path, encoding="utf-8", low_memory=False) if file_path.endswith(
            ".csv") else pd.read_excel(file_path, sheet_name=0)

        if df.empty:
            messagebox.showerror("Error", "Loaded file is empty or could not be read.")
            return None

        update_columns()
        display_data(df)
        messagebox.showinfo("Success", "File uploaded successfully!")
        return df  # Return the loaded DataFrame

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")
        print("Upload Error:", e)
        return None


def display_data(data, tree, sort_orders, toggle_sort_order):
    """Displays the loaded DataFrame in the Tkinter Treeview."""
    if data is None:
        return

    tree.delete(*tree.get_children())  # Clear existing data
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"  # Ensure only table headers are visible, not row indices

    for col in data.columns:
        arrow = " ⬆" if sort_orders.get(col) else " ⬇" if col in sort_orders else ""
        tree.heading(col, text=f"{col}{arrow}", command=lambda c=col: toggle_sort_order(c))
        tree.column(col, width=150, anchor="center")

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

    tree.update_idletasks()
