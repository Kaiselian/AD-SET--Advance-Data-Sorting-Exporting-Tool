import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb

class ExportWidgets:
    def __init__(self, parent, filtered_df, export_data_callback):
        self.parent = parent
        self.filtered_df = filtered_df
        self.export_data_callback = export_data_callback
        self.create_widgets()

    def create_widgets(self):
        export_menu_btn = tb.Menubutton(self.parent, text="📤 Export", bootstyle="warning")
        export_menu_btn.pack(side=tk.RIGHT, padx=10)

        export_menu = tk.Menu(export_menu_btn, tearoff=0)
        export_menu.add_command(label="📤 Export as CSV", command=lambda: self.export_data_callback("csv"))
        export_menu.add_command(label="📤 Export as Excel", command=lambda: self.export_data_callback("xlsx"))
        export_menu.add_command(label="📤 Export as PDF", command=lambda: self.export_data_callback("pdf"))

        export_menu_btn["menu"] = export_menu