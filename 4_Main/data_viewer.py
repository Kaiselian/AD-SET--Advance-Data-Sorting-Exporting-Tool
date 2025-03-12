import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
import pandas as pd

def view_data_table(root, data):
    """
    Creates a new window to display data in a scrollable treeview table.

    :param root: The main Tkinter root window.
    :param data: Pandas DataFrame containing the data.
    """
    if data is None or data.empty:
        tb.messagebox.showerror("Error", "No data to display!")
        return

    # Create a new top-level window
    data_window = tk.Toplevel(root)
    data_window.title("Data Viewer")
    data_window.geometry("800x500")

    # Create a frame for the treeview with scrollbars
    frame = tb.Frame(data_window)
    frame.pack(pady=10, fill=tk.BOTH, expand=True)

    # Create a Treeview widget
    tree = ttk.Treeview(frame, show="headings", selectmode="browse")

    # Add vertical and horizontal scrollbars
    tree_scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree_scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

    tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(fill=tk.BOTH, expand=True)

    # Define columns and headings
    tree["columns"] = list(data.columns)
    for col in data.columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, anchor="center", width=150)

    # Insert data rows
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

    # Pack the treeview
    tree.pack(fill=tk.BOTH, expand=True)
