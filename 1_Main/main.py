from tkinter import Tk
import ttkbootstrap as tb
from toolbar import create_top_toolbar
from pdf_loader import load_pdf
#from some_other_modules import upload_file, search_and_generate, clear_filters, load_pdf, convert_pdf_to_excel, export_filtered_data, export_each_row_as_pdf, export_filled_pdfs
from data_filter import clear_filters, search_and_generate
from data_loader import upload_file, display_data

# 🟢 Create Root Window
root = Tk()
root.title("PDF Form Filler")

# 🟢 Create Top Toolbar
top_frame = create_top_toolbar(root, upload_file, search_and_generate, clear_filters, load_pdf, convert_pdf_to_excel, export_filtered_data, export_each_row_as_pdf, export_filled_pdfs)

# 🟢 Load PDF Button
btn_load_pdf = tb.Button(root, text="📂 Load PDF", bootstyle="primary", command=lambda: load_pdf(root, export_filled_pdfs, extract_to_excel, set_extraction_start, update_columns))
btn_load_pdf.pack(pady=10)

# Store search variables in a dictionary
search_vars = {
    "search_var": search_var,
    "sub_search_var": sub_search_var,
    "column_var": column_var,
    "sub_search_column_var": sub_search_column_var,
    "filter_var": filter_var,
}
# Upload Button
upload_btn = tb.Button(root, text="📂 Upload File", bootstyle="primary",
                       command=lambda: upload_file(update_columns, lambda df: display_data(df, tree, sort_orders, toggle_sort_order)))
upload_btn.pack(side=tk.LEFT, padx=10)


# Connect buttons to the new functions
clear_btn = tb.Button(root, text="❌ Clear Filters", bootstyle="danger",
                      command=lambda: clear_filters(df, search_vars, sort_orders, display_data))
clear_btn.pack(side=tk.LEFT, padx=10)

search_btn = tb.Button(root, text="🔍", bootstyle="success",
                       command=lambda: search_and_generate(df, search_vars, display_data))
search_btn.pack(side=tk.LEFT, padx=10)

root.mainloop()
