# toolbar.py
import tkinter as tk
import ttkbootstrap as tb
from tkinter import ttk


def create_top_toolbar(root, upload_file, search_and_generate, clear_filters, load_pdf, convert_pdf_to_excel,
                       export_filtered_data, export_each_row_as_pdf, export_filled_pdfs):
    """Creates and returns the top toolbar frame."""

    # 🟢 UI Layout - Top Bar
    top_frame = tb.Frame(root)
    top_frame.pack(pady=10, fill=tk.X, padx=20)

    # 📂 Upload Button
    upload_btn = tb.Button(top_frame, text="📂 Upload File", bootstyle="primary", command=upload_file)
    upload_btn.pack(side=tk.LEFT, padx=10)

    # 🔍 Search Bar
    search_var = tk.StringVar()
    search_entry = tb.Entry(top_frame, textvariable=search_var, width=40)
    search_entry.pack(side=tk.LEFT, padx=10)
    search_entry.bind("<Return>", lambda event: search_and_generate())  # ENTER triggers search

    search_btn = tb.Button(top_frame, text="🔍", bootstyle="success", command=search_and_generate)
    search_btn.pack(side=tk.LEFT, padx=10)

    # 🔍 Sub-Search Bar & Column Selection
    sub_search_var = tk.StringVar()
    sub_search_entry = tb.Entry(top_frame, textvariable=sub_search_var, width=40)
    sub_search_entry.pack(side=tk.LEFT, padx=10)
    sub_search_entry.bind("<Return>", lambda event: search_and_generate())  # ENTER triggers sub-search

    sub_search_column_var = tk.StringVar(value="All Columns")
    sub_search_column_dropdown = ttk.Combobox(top_frame, textvariable=sub_search_column_var, state="readonly")
    sub_search_column_dropdown.pack(side=tk.LEFT, padx=10)

    sub_search_btn = tb.Button(top_frame, text="🔍 Sub-Search", bootstyle="success", command=search_and_generate)
    sub_search_btn.pack(side=tk.LEFT, padx=10)

    # 🔽 Column Dropdown
    column_var = tk.StringVar(value="All Columns")
    column_dropdown = ttk.Combobox(top_frame, textvariable=column_var, state="readonly")

    # 🔍 Filter Type Dropdown
    filter_var = tk.StringVar(value="Contains")
    filter_dropdown = ttk.Combobox(top_frame, textvariable=filter_var, state="readonly",
                                   values=["Contains", "Equals", "Starts with"])

    # ❌ Clear Button
    clear_btn = tb.Button(top_frame, text="❌ Clear Filters", bootstyle="danger", command=clear_filters)
    clear_btn.pack(side=tk.LEFT, padx=10)

    # 📂 Load PDF for Preview
    btn_load_pdf = ttk.Button(top_frame, text="📂 Load PDF", command=load_pdf)
    btn_load_pdf.pack(side=tk.LEFT, padx=5)

    # 📥 PDF TO EXCEL Button
    pdf_to_excel_btn = tb.Button(top_frame, text="📥 PDF to Excel", bootstyle="info", command=convert_pdf_to_excel)
    pdf_to_excel_btn.pack(side=tk.RIGHT, padx=10)

    # 📤 Export Menu Button
    export_menu_btn = tb.Menubutton(top_frame, text="📤 Export", bootstyle="warning")
    export_menu_btn.pack(side=tk.RIGHT, padx=10)

    # 🔽 Dropdown Menu for Export
    export_menu = tk.Menu(export_menu_btn, tearoff=0)
    export_menu.add_command(label="📤 Export as CSV", command=lambda: export_filtered_data("csv"))
    export_menu.add_command(label="📤 Export as Excel", command=lambda: export_filtered_data("xlsx"))
    export_menu.add_command(label="📤 Export Full PDF", command=lambda: export_filtered_data("pdf"))
    export_menu.add_command(label="📤 Export Individual PDFs", command=export_each_row_as_pdf)
    export_menu_btn["menu"] = export_menu

    # 📤 Export PDFs Button
    btn_export_pdf = tb.Button(top_frame, text="📤 Export PDFs", bootstyle="success", command=export_filled_pdfs)
    btn_export_pdf.pack(side=tk.RIGHT, padx=10)

    return top_frame
