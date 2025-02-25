import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
import ttkbootstrap as tb  # Modern UI Framework
import darkdetect #detect system default active UI
from fpdf import FPDF
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import openpyxl
import pdfplumber  # Extract tables from PDFs
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import subprocess
import sys
import os


# Detect System Theme (Light/Dark)
def get_system_theme():
    return "darkly" if darkdetect.isDark() else "journal"

# 🔄 Function to Change Theme
def change_theme(selected_theme):
    global root
    root.style.theme_use(selected_theme)  # Apply the new theme instantly

# Global variables
df = None
pdf_path = None
pdf_document = None
pdf_img = None
pdf_canvas = None
text_boxes = []
box_data = []
base_name = "Invoice"  # Default base name
current_zoom = 1.0  # Zoom Level
selected_font = tk.StringVar(value="Arial")  # Default font
selected_size = tk.IntVar(value=12)  # Default size

# Global variable to store the last filtered dataset
filtered_df = None

# Initialize GUI
theme = "darkly" if darkdetect.isDark() else "journal"

root = tb.Window(themename=theme)  # Default theme, fixed
root.title("Advanced Data Search & Export Tool 1.13.5.7")
root.geometry("1920x1080")
root.state("zoomed")

# 🟢 **Load PDF Function**
def load_pdf():
    global pdf_path, pdf_document, pdf_canvas, pdf_img, text_boxes, box_data, current_zoom

    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    pdf_document = fitz.open(pdf_path)
    page = pdf_document[0]
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.thumbnail((800, 1000))  # Adjust initial size
    pdf_img = ImageTk.PhotoImage(img)

    # 🆕 **Open in New Window**
    pdf_window = tk.Toplevel(root)
    pdf_window.title("PDF Preview - Assign Data Fields")
    pdf_window.geometry("1200x900")
    pdf_window.state("zoomed")

    # 🟢 **Frame for Side Controls**
    frame_right = tk.Frame(pdf_window, width=300, bg="#f0f0f0")
    frame_right.pack(side=tk.RIGHT, fill=tk.Y)

    # ✅ Export Button
    btn_export = tk.Button(frame_right, text="📤 Export PDFs", command=export_filled_pdfs)
    btn_export.pack(pady=10, padx=10, fill=tk.X)

    # ✅ Zoom Controls
    def zoom_in():
        global current_zoom
        current_zoom *= 1.2
        update_preview()

    def zoom_out():
        global current_zoom
        current_zoom /= 1.2
        update_preview()

    btn_zoom_in = tk.Button(frame_right, text="➕ Zoom In", command=zoom_in)
    btn_zoom_in.pack(pady=5, padx=10, fill=tk.X)

    btn_zoom_out = tk.Button(frame_right, text="➖ Zoom Out", command=zoom_out)
    btn_zoom_out.pack(pady=5, padx=10, fill=tk.X)

    # 🟢 **Canvas for PDF Preview**
    pdf_canvas = tk.Canvas(pdf_window, width=img.width, height=img.height)
    pdf_canvas.pack(side=tk.LEFT, expand=True)
    pdf_canvas.create_image(0, 0, anchor=tk.NW, image=pdf_img)

    # ✅ **Update PDF Preview (After Zoom)**
    def update_preview():
        page = pdf_document[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(current_zoom, current_zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        pdf_img_resized = ImageTk.PhotoImage(img)
        pdf_canvas.create_image(0, 0, anchor=tk.NW, image=pdf_img_resized)
        pdf_canvas.image = pdf_img_resized

    # 🆕 **Function to Add Text Box**
    def add_text_box():
        if df is None:
            messagebox.showerror("Error", "Please upload an Excel/CSV file first.")
            return

        # ✅ **Frame for Resizable & Draggable Text Box**
        frame = tk.Frame(pdf_canvas, bd=1, relief="ridge", bg="white")
        frame.place(x=100, y=100, width=150, height=50)

        # ✅ **Multi-line Text Widget Inside Frame**
        entry = tk.Text(frame, font=(selected_font.get(), selected_size.get()), wrap="word", height=2, width=20)
        entry.pack(fill=tk.BOTH, expand=True)

        # ✅ Store Text Box Info
        text_boxes.append({"frame": frame, "entry": entry, "column": None})

        # ✅ Dropdown for Data Column
        col_dropdown = ttk.Combobox(frame_right, values=list(df.columns), state="readonly")
        col_dropdown.pack(pady=2, padx=10, fill=tk.X)
        text_boxes[-1]["column"] = col_dropdown

        # 🟢 **Dragging Functionality**
        def on_drag(event):
            x = frame.winfo_x() + (event.x - frame._drag_start_x)
            y = frame.winfo_y() + (event.y - frame._drag_start_y)
            frame.place(x=x, y=y)

        def start_drag(event):
            frame._drag_start_x = event.x
            frame._drag_start_y = event.y

        frame.bind("<ButtonPress-1>", start_drag)
        frame.bind("<B1-Motion>", on_drag)

        # 🟢 **Resizable Text Box**
        def start_resize(event):
            frame._resize_start_x = event.x
            frame._resize_start_y = event.y

        def resize(event):
            new_width = max(50, frame.winfo_width() + (event.x - frame._resize_start_x))
            new_height = max(30, frame.winfo_height() + (event.y - frame._resize_start_y))
            frame.place(width=new_width, height=new_height)
            entry.config(width=int(new_width / 10), height=int(new_height / 20))

        resize_handle = tk.Label(frame, text="⤡", cursor="bottom_right_corner", bg="lightgray")
        resize_handle.place(relx=1.0, rely=1.0, anchor="se")

        resize_handle.bind("<ButtonPress-1>", start_resize)
        resize_handle.bind("<B1-Motion>", resize)

        # ✅ Button to Add Text Box

    btn_add_box = tk.Button(frame_right, text="➕ Add Text Box", command=add_text_box)
    btn_add_box.pack(pady=10, padx=10, fill=tk.X)

def adjust_font():
    selected_font = font_var.get()
    selected_size = size_var.get()
    entry.config(font=(selected_font, selected_size))

    font_var = tk.StringVar(value="Arial")
    size_var = tk.IntVar(value=12)

    font_dropdown = ttk.Combobox(frame_right, textvariable=font_var, values=["Arial", "Times New Roman", "Courier"], state="readonly")
    font_dropdown.pack(pady=2, padx=10, fill=tk.X)
    size_dropdown = ttk.Combobox(frame_right, textvariable=size_var, values=[8, 10, 12, 14, 16, 18, 20], state="readonly")
    size_dropdown.pack(pady=2, padx=10, fill=tk.X)
    btn_apply_font = tk.Button(frame_right, text="Apply Font", command=adjust_font)
    btn_apply_font.pack(pady=5, padx=10, fill=tk.X)

    btn_add_box = tk.Button(frame_right, text="➕ Add Text Box", command=add_text_box)
    btn_add_box.pack(pady=10, padx=10, fill=tk.X)


def set_extraction_start():
    global segment_start
    segment_start = simpledialog.askstring("Set Extraction Start", "Enter keyword or phrase to start extraction from:")
    if segment_start:
        messagebox.showinfo("Success", f"Extraction will start from: {segment_start}")


def extract_to_excel():
    if pdf_document is None:
        messagebox.showerror("Error", "No PDF loaded!")
        return

    extracted_data = []
    for page_num, page in enumerate(pdf_document):
        text_blocks = page.get_text("blocks")  # Extract text in structured segments
        filtered_blocks = []
        start_found = False if segment_start else True

        for block in text_blocks:
            if segment_start and segment_start in block[4]:
                start_found = True
            if start_found:
                filtered_blocks.append(block[4])

        page_data = {f"Segment {i + 1}": text for i, text in enumerate(filtered_blocks)}
        page_data["Page Number"] = page_num + 1
        extracted_data.append(page_data)

    df = pd.DataFrame(extracted_data)

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"Data extracted and saved to {save_path}")


def set_custom_name():
    global base_name
    new_name = simpledialog.askstring("Custom Name", "Enter base name for exported PDFs:")
    if new_name:
        base_name = new_name

def export_filled_pdfs():
    if df is None or df.empty:
        messagebox.showerror("Error", "No data file uploaded!")
        return
    if pdf_document is None:
        messagebox.showerror("Error", "No PDF template loaded!")
        return
    if not text_boxes:
        messagebox.showerror("Error", "No text fields assigned for data mapping!")
        return

    save_folder = filedialog.askdirectory()
    if not save_folder:
        return

    # ✅ **Ensure Text Boxes Have Assigned Columns**
    for box in text_boxes:
        if not box["column"].get():
            messagebox.showerror("Error", "Some text boxes have no assigned columns!")
            return

    # ✅ **Generate PDFs for Each Row**
    for index, row in df.iterrows():
        filled_pdf = fitz.open()
        for page in pdf_document:
            new_page = filled_pdf.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, pdf_document, page.number)

            for box in text_boxes:
                field_column = box["column"].get()
                if field_column in df.columns:
                    text_value = str(row[field_column])
                    x, y = box["frame"].winfo_x(), box["frame"].winfo_y()
                    new_page.insert_text((x, y), text_value, fontsize=selected_size.get(), color=(0, 0, 0))

        output_file = os.path.join(save_folder, f"{base_name}_{index + 1}.pdf")
        filled_pdf.save(output_file)
        filled_pdf.close()

    messagebox.showinfo("Success", f"PDFs saved in {save_folder}")


# 🟢 Upload File Function
def upload_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False) if file_path.endswith(".csv") else pd.read_excel(file_path, sheet_name=0)

            if df.empty:
                messagebox.showerror("Error", "Loaded file is empty or could not be read.")
                return

            update_columns()
            display_data(df)  # Now tree exists, so no error.
            messagebox.showinfo("Success", "File uploaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            print("Upload Error:", e)  # Debugging

# Display Data in Treeview with Proper Table Formatting
def display_data(data):
    global filtered_df
    filtered_df = data  # Store filtered data

    tree.delete(*tree.get_children())  # Clear existing data
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"  # Ensure only table headers are visible, not row indices

    # Adjust columns dynamically
    for col in data.columns:
        arrow = ""
        if col in sort_orders and sort_orders[col] is not None:  # Show arrow only if column was sorted
            arrow = " ⬆" if sort_orders[col] else " ⬇"

        tree.heading(col, text=f"{col}{arrow}", command=lambda c=col: toggle_sort_order(c))
        tree.column(col, width=150, anchor="center")  # Set a default width

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

    tree.update_idletasks()  # Refresh to apply changes

# 🔼🔽 Toggle Sort Order
sort_orders = {}  # Dictionary to track column sorting order
def toggle_sort_order(column):
    global filtered_df

    if filtered_df is None or column not in filtered_df.columns:
        messagebox.showerror("Error", "Invalid column selection.")
        return

    # Reset all sorting indicators except the selected column
    for col in sort_orders:
        if col != column:
            sort_orders[col] = None  # Remove sorting order for other columns

    # Toggle sorting order for the selected column
    sort_orders[column] = not sort_orders.get(column, False)
    ascending = sort_orders[column]

    # Sort the data based on the selected column and order
    filtered_df = filtered_df.sort_values(by=column, ascending=ascending)

    display_data(filtered_df)  # Refresh sorted data


# Clear Data Searched and Reset Sorting
def clear_filters():
    global filtered_df, sort_orders
    if df is None:
        messagebox.showerror("Error", "No data loaded to clear filters.")
        return

    filtered_df = df.copy()  # Reset data
    sort_orders = {}  # Reset sorting order

    search_var.set("")
    sub_search_var.set("")
    column_var.set("All Columns")
    sub_search_column_var.set("All Columns")
    filter_var.set("Contains")

    display_data(filtered_df)  # Refresh without sorting icons


# 🔍 Combined Search (Main & Sub-Search)
def search_and_generate():
    global df, filtered_df

    if df is None:
        messagebox.showerror("Error", "Please upload a file first.")
        return

    main_query = search_var.get().strip()
    sub_query = sub_search_var.get().strip()
    main_column = column_var.get()
    sub_column = sub_search_column_var.get()
    filter_type = filter_var.get()

    if not main_query and not sub_query:
        messagebox.showerror("Error", "Please enter a search term.")
        return

    filtered_data = df.copy()

    # 🔹 Apply Main Search
    if main_query:
        if main_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(main_query, case=False, na=False).any(), axis=1)
            ]
        else:
            if filter_type == "Contains":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.contains(main_query, case=False, na=False)]
            elif filter_type == "Equals":
                filtered_data = filtered_data[filtered_data[main_column].astype(str) == main_query]
            elif filter_type == "Starts with":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.startswith(main_query, na=False)]

    # 🔹 Apply Sub-Search on Filtered Data
    if sub_query:
        if sub_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(sub_query, case=False, na=False).any(), axis=1)
            ]
        else:
            filtered_data = filtered_data[filtered_data[sub_column].astype(str).str.contains(sub_query, case=False, na=False)]

    # 🛑 FIXED: Correct placement of "No Results" message
    if filtered_data.empty:
        messagebox.showinfo("No Results", "No matching records found.")
        return

    display_data(filtered_data)  # ✅ Display only once
    filtered_df = filtered_data  # ✅ Store for sorting

    # Store filtered data for sorting

    if filtered_data.empty:
        messagebox.showinfo("No Results", "No matching records found.")
        return

    display_data(filtered_data)

# 🟢 **Function to Export PDFs**
def export_filled_pdfs():
    if df is None or df.empty:
        messagebox.showerror("Error", "No data file uploaded!")
        return
    if pdf_document is None:
        messagebox.showerror("Error", "No PDF template loaded!")
        return
    if not text_boxes:
        messagebox.showerror("Error", "No text fields assigned for data mapping!")
        return

    save_folder = filedialog.askdirectory()
    if not save_folder:
        return

    # ✅ **Ensure Text Boxes Have Assigned Columns**
    for box in text_boxes:
        if not box["column"].get():
            messagebox.showerror("Error", "Some text boxes have no assigned columns!")
            return

    # ✅ **Generate PDFs for Each Row**
    for index, row in df.iterrows():
        filled_pdf = fitz.open()
        for page in pdf_document:
            new_page = filled_pdf.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, pdf_document, page.number)

            for box in text_boxes:
                field_column = box["column"].get()
                if field_column in df.columns:
                    text_value = str(row[field_column])
                    x, y = box["frame"].winfo_x(), box["frame"].winfo_y()
                    new_page.insert_text((x, y), text_value, fontsize=selected_size.get(), color=(0, 0, 0))

        output_file = os.path.join(save_folder, f"{base_name}_{index + 1}.pdf")
        filled_pdf.save(output_file)
        filled_pdf.close()

    messagebox.showinfo("Success", f"PDFs saved in {save_folder}")

# 🖨 Convert Excel DataFrame to PDF
def save_df_as_pdf(df, save_path):
    # Define Portrait A4 size as default
    page_width, page_height = landscape(A4)  # Default to Portrait A4
    column_count = len(df.columns)

    # Estimate column widths dynamically (evenly distribute within the page width)
    max_col_width = (page_width - 40) / column_count  # Subtract margins and divide by columns
    column_widths = [max_col_width] * column_count  # Apply the width to all columns

    # If total column width exceeds page width, switch to Landscape mode
    if sum(column_widths) > page_width - 40:
        page_width, page_height = landscape(A4)
        max_col_width = (page_width - 40) / column_count
        column_widths = [max_col_width] * column_count

    # Create the PDF document
    doc = SimpleDocTemplate(save_path, pagesize=(page_width, page_height))
    elements = []

    # Convert DataFrame to a list of lists
    data = [df.columns.tolist()] + df.astype(str).values.tolist()

    # Create a table with adjusted column widths
    table = Table(data, colWidths=column_widths)

    # Add Styling
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center align all cells
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for header
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),  # Padding for header
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Row background
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Table grid
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # Reduce font size to fit data
    ])

    table.setStyle(style)

    elements.append(table)
    doc.build(elements)

    print(f"PDF saved successfully in {'Landscape' if page_width > page_height else 'Portrait'} orientation!")

    # Scale the table to fit within the printable area
    available_width = page_width - 40  # Margins (20 on each side)
    available_height = page_height - 40  # Margins (20 on top/bottom)

    # Get the natural table size
    table_width, table_height = table.wrap(available_width, available_height)

    # Calculate scaling factor based on the available space
    scale_factor = min(available_width / table_width if table_width else 1,
                       available_height / table_height if table_height else 1)

    # Apply scaling factor to columns and rows
    scaled_column_widths = [w * scale_factor if w else 0 for w in column_widths]
    table._colWidths = scaled_column_widths
    table._rowHeights = [h * scale_factor if h else 0 for h in table._rowHeights]

    # Recalculate table size after scaling
    table_width, table_height = table.wrap(available_width, available_height)

    # Ensure table still fits within page, reapply scale if necessary
    if table_width > available_width or table_height > available_height:
        print("Warning: The table still doesn't fit despite scaling. Check column widths or data size.")

    elements.append(table)

    # Build the PDF document
    doc.build(elements)

    print(f"PDF saved successfully with {'Landscape' if page_width > page_height else 'Portrait'} orientation!")

 # Function to Export Each Row Individually as PDF
def export_each_row_as_pdf():
    global filtered_df
    if filtered_df is None or filtered_df.empty:
        messagebox.showerror("Error", "No data to export.")
        return

    save_directory = filedialog.askdirectory()
    if not save_directory:
        return

    styles = getSampleStyleSheet()
    for index, row in filtered_df.iterrows():
        file_path = os.path.join(save_directory, f"Row_{index + 1}.pdf")
        doc = SimpleDocTemplate(file_path, pagesize=A4)
        elements = []

        elements.append(Paragraph("Row Data", styles['Title']))
        table_data = [[col, str(row[col])] for col in filtered_df.columns]
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))

        elements.append(table)
        doc.build(elements)

    messagebox.showinfo("Success", "Each row exported as an individual PDF.")

#PDF TO EXCEL
def convert_pdf_to_excel():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return

    extracted_data = []
    headers = None  # Store column headers separately

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                if table:
                    first_row = table[0]  # First row of the table
                    if headers is None:  # Set headers only once
                        #headers = first_row
                        extracted_data.append(headers)
                    else:
                        # If the first row is just numbers, ignore it
                        if all(cell.isdigit() for cell in first_row if cell):
                            table = table[1:]  # Skip first row

                    for row in table:
                        extracted_data.append(row)

    if extracted_data:
        df = pd.DataFrame(extracted_data[1:], columns=extracted_data[0])  # Use first row as headers
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", "PDF converted to Excel successfully!")
    else:
        messagebox.showerror("Error", "No tables found in the PDF.")


# 🔹 UI Layout - Top Bar
top_frame = tb.Frame(root)
top_frame.pack(pady=10, fill=tk.X, padx=20)

upload_btn = tb.Button(top_frame, text="📂 Upload File", bootstyle="primary", command=upload_file)
upload_btn.pack(side=tk.LEFT, padx=10)

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
filter_dropdown = ttk.Combobox(top_frame, textvariable=filter_var, state="readonly", values=["Contains", "Equals", "Starts with"])

# Clear Button
clear_btn = tb.Button(top_frame, text="❌ Clear Filters", bootstyle="danger", command=clear_filters)
clear_btn.pack(side=tk.LEFT, padx=10)

# Load PDF for Preview
btn_load_pdf = ttk.Button(top_frame, text="📂 Load PDF", command=load_pdf)
btn_load_pdf.pack(side=tk.LEFT, padx=5)

# PDF TO EXCEL Button
pdf_to_excel_btn = tb.Button(top_frame, text="📥 PDF to Excel", bootstyle="info", command=convert_pdf_to_excel)
pdf_to_excel_btn.pack(side=tk.RIGHT, padx=10)

# Create a Menu Button for Export Options
export_menu_btn = tb.Menubutton(top_frame, text="📤 Export", bootstyle="warning")
export_menu_btn.pack(side=tk.RIGHT, padx=10)

# Create the Dropdown Menu
export_menu = tk.Menu(export_menu_btn, tearoff=0)
export_menu.add_command(label="📤 Export as CSV", command=lambda: export_filtered_data("csv"))
export_menu.add_command(label="📤 Export as Excel", command=lambda: export_filtered_data("xlsx"))
export_menu.add_command(label="📤 Export Full PDF", command=lambda: export_filtered_data("pdf"))
export_menu.add_command(label="📤 Export Individual PDFs", command=lambda: export_each_row_as_pdf())

# Attach the Menu to the Button
export_menu_btn["menu"] = export_menu

# 🟢 **Add Export Button**
btn_export_pdf = tb.Button(top_frame, text="📤 Export PDFs", bootstyle="success", command=export_filled_pdfs)
btn_export_pdf.pack(side=tk.RIGHT, padx=10)

# Create a frame for the Treeview with Scrollbars
frame2 = tb.Frame(root)
frame2.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

# Add a Vertical Scrollbar
tree_scroll_y = tk.Scrollbar(frame2, orient="vertical")
tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

# Add a Horizontal Scrollbar
tree_scroll_x = tk.Scrollbar(frame2, orient="horizontal")
tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

# Create the Treeview
tree = ttk.Treeview(frame2, style="Custom.Treeview", yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree.pack(pady=10, fill=tk.BOTH, expand=True)

# Configure the Scrollbars to Scroll the Treeview
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

# 🔄 Update column dropdown when a file is loaded
def update_columns():
    if df is not None:
        column_dropdown["values"] = ["All Columns"] + list(df.columns)
        column_var.set("All Columns")
        sub_search_column_dropdown["values"] = ["All Columns"] + list(df.columns)
        sub_search_column_var.set("All Columns")

# New Window - for different processing
def new_window():
    try:
        subprocess.Popen([sys.executable, sys.argv[0]])  # Runs a new instance of the current script
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open new window: {e}")

# Create the main menu
menu_bar = tk.Menu(root)

# File Menu
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Upload File", command=upload_file)

# Export as a sub-menu of File
export_menu = tk.Menu(file_menu, tearoff=0)
export_menu.add_command(label="Export as CSV", command=lambda: export_filtered_data("csv"))
export_menu.add_command(label="Export as Excel", command=lambda: export_filtered_data("xlsx"))
export_menu.add_command(label="Export as PDF", command=lambda: export_filtered_data("pdf"))
file_menu.add_cascade(label="Export", menu=export_menu)
menu_bar.add_cascade(label="File", menu=file_menu)

# New Windows Button
new_window_menu = tk.Menu(menu_bar, tearoff=0)
new_window_menu.add_command(label="Open New Window", command=new_window)
menu_bar.add_cascade(label="New Window", menu=new_window_menu)  # Correct usage


# Theme Menu
theme_menu = tk.Menu(menu_bar, tearoff=0)
theme_options = {
    "darkly": "🌙",
    "journal": "📖",
    "flatly": "📄",
    "cyborg": "🤖",
    "superhero": "🦸",
    "minty": "🌿"}

for theme, emoji in theme_options.items():
    theme_menu.add_command(label=f"{emoji} {theme}", command=lambda t=theme: change_theme(t))

menu_bar.add_cascade(label="Theme", menu=theme_menu)

# Add the menu bar to the root window
root.config(menu=menu_bar)

root.mainloop()
