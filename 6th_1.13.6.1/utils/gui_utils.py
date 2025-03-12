import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb

def create_treeview(parent):
    """
    Creates a Treeview widget with scrollbars.
    """
    # Create a frame for the Treeview and scrollbars
    frame = tb.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)

    # Vertical Scrollbar
    tree_scroll_y = tk.Scrollbar(frame, orient="vertical")
    tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

    # Horizontal Scrollbar
    tree_scroll_x = tk.Scrollbar(frame, orient="horizontal")
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    # Create the Treeview
    tree = ttk.Treeview(frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
    tree.pack(fill=tk.BOTH, expand=True)

    # Configure the Scrollbars
    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)

    return tree, frame

def display_data(tree, data, sort_orders=None):
    """
    Displays data in the Treeview widget.
    """
    if sort_orders is None:
        sort_orders = {}

    # Clear existing data
    tree.delete(*tree.get_children())

    # Set up columns
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"  # Ensure only table headers are visible, not row indices

    # Adjust columns dynamically
    for col in data.columns:
        arrow = ""
        if col in sort_orders and sort_orders[col] is not None:  # Show arrow only if column was sorted
            arrow = " ⬆" if sort_orders[col] else " ⬇"

        tree.heading(col, text=f"{col}{arrow}", command=lambda c=col: toggle_sort_order(c))
        tree.column(col, width=150, anchor="center")  # Set a default width

    # Insert data rows
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

    tree.update_idletasks()  # Refresh to apply changes