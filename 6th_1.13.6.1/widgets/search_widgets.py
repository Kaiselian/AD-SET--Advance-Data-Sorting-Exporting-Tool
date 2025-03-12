import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb

class SearchWidgets:
    def __init__(self, parent, tree, display_data_callback):
        self.parent = parent
        self.tree = tree
        self.display_data_callback = display_data_callback
        self.search_var = tk.StringVar()
        self.sub_search_var = tk.StringVar()
        self.column_var = tk.StringVar(value="All Columns")
        self.sub_search_column_var = tk.StringVar(value="All Columns")
        self.filter_var = tk.StringVar(value="Contains")
        self.create_widgets()

    def create_widgets(self):
        self.search_entry = tb.Entry(self.parent, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=10)
        self.sub_search_entry = tb.Entry(self.parent, textvariable=self.sub_search_var, width=40)
        self.sub_search_entry.pack(side=tk.LEFT, padx=10)
        self.column_dropdown = ttk.Combobox(self.parent, textvariable=self.column_var, state="readonly")
        self.column_dropdown.pack(side=tk.LEFT, padx=10)
        self.sub_search_column_dropdown = ttk.Combobox(self.parent, textvariable=self.sub_search)