import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import ttkbootstrap as tb
from widgets.search_widgets import SearchWidgets
from utils.file_utils import upload_file, export_filtered_data
from utils.data_utils import filter_data, display_data
from utils.docx_filler import fill_docx_template
from views.pdf_view import PDFView


# Global variables
df = None
pdf_path = None
pdf_document = None
pdf_img = None
pdf_canvas = None

class MainView(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.df = None
        self.filtered_df = None
        self.sort_orders = {}  # Track sorting order for columns
        self.create_widgets()

    def clear_filters(self):
        """Resets all search filters and refreshes the dataset."""
        if self.df is None:
            messagebox.showerror("Error", "No data loaded to clear filters.")
            return

        self.filtered_df = self.df.copy()  # Reset data
        self.sort_orders = {}  # Reset sorting order

        self.search_var.set("")
        self.sub_search_var.set("")
        self.column_var.set("All Columns")
        self.sub_search_column_var.set("All Columns")
        self.filter_var.set("Contains")

        self.display_data(self.search_var.get(), self.sub_search_var.get(), self.column_var.get(),
                          self.sub_search_column_var.get(), self.filter_var.get())  # Refresh without sorting icons

    def upload_file(self):
        """Handles the file upload."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if file_path:
            try:
                #load your data.
                #example loading of data.
                import pandas as pd
                self.df = pd.read_excel(file_path) #load excel file.
                self.search_widgets.column_dropdown["values"] = ["All Columns"] + list(self.df.columns)
            except AttributeError:
                print("dataframe not initialized.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to upload file: {e}")
        else:
            print("No file selected.")

    def display_data(self, search_query="", sub_query="", main_column="All Columns", sub_column="All Columns", filter_type="Contains"):
        """Filters and updates the Treeview based on search criteria."""
        if self.df is None:
            return

        # Apply filtering
        filtered_df = filter_data(self.df, search_query, sub_query, main_column, sub_column, filter_type)

        # Update Treeview
        display_data(self.tree, filtered_df, self.sort_orders)

    def open_pdf_view(self):
        """Opens the PDF view with the loaded PDF."""
        if self.pdf_path:
            pdf_window = tk.Toplevel(self.parent)
            pdf_window.title("PDF Preview")
            pdf_view_instance = PDFView(pdf_window, self.pdf_path)  # create instance
            pdf_view_instance.pack(fill=tk.BOTH, expand=True)  # pack the instance
        else:
            messagebox.showerror("Error", "No PDF loaded.")

    def start_processing(self):
        """Handles the processing of data and template files."""
        if not self.df or not self.template_file or not self.output_folder:
            messagebox.showerror("Error", "Please upload all required files!")
            return

        # Fill DOCX Templates for Each Row
        filled_files = fill_docx_template(self.template_file, self.df, self.output_folder, file_prefix="invoice")
        if not filled_files:
            messagebox.showerror("Error", "Failed to fill the DOCX template.")
            return

        messagebox.showinfo("Success", f"{len(filled_files)} documents saved in {self.output_folder}")

    def open_pdf(self):
        """Opens a PDF file in the PDF preview window."""
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if pdf_path:
            PDFView(self, pdf_path)  # Open the PDF preview window

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

    def create_widgets(self):

        # 🔹 UI Layout - Top Bar
        top_frame = tb.Frame(self)  # change root to self.
        top_frame.pack(pady=10, fill=tk.X, padx=20)

        upload_btn = tb.Button(top_frame, text="📂 Upload File", bootstyle="primary", command=self.upload_file)  # add self.
        upload_btn.pack(side=tk.LEFT, padx=10)

        search_var = tk.StringVar()
        search_entry = tb.Entry(top_frame, textvariable=search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=10)
        search_entry.bind("<Return>", lambda event: self.search_and_generate())  # ENTER triggers search #add self.

        search_btn = tb.Button(top_frame, text="🔍", bootstyle="success", command=self.search_and_generate)  # add self.
        search_btn.pack(side=tk.LEFT, padx=10)

        # 🔍 Sub-Search Bar & Column Selection
        sub_search_var = tk.StringVar()
        sub_search_entry = tb.Entry(top_frame, textvariable=sub_search_var, width=40)
        sub_search_entry.pack(side=tk.LEFT, padx=10)
        sub_search_entry.bind("<Return>", lambda event: self.search_and_generate())  # ENTER triggers sub-search #add self.

        sub_search_column_var = tk.StringVar(value="All Columns")
        sub_search_column_dropdown = ttk.Combobox(top_frame, textvariable=sub_search_column_var, state="readonly")
        sub_search_column_dropdown.pack(side=tk.LEFT, padx=10)

        sub_search_btn = tb.Button(top_frame, text="🔍 Sub-Search", bootstyle="success", command=self.search_and_generate)  # add self.
        sub_search_btn.pack(side=tk.LEFT, padx=10)

        # 🔽 Column Dropdown
        column_var = tk.StringVar(value="All Columns")
        column_dropdown = ttk.Combobox(top_frame, textvariable=column_var, state="readonly")

        # 🔍 Filter Type Dropdown
        filter_var = tk.StringVar(value="Contains")
        filter_dropdown = ttk.Combobox(top_frame, textvariable=filter_var, state="readonly", values=["Contains", "Equals", "Starts with"])

        # Clear Button
        clear_btn = tb.Button(top_frame, text="❌ Clear Filters", bootstyle="danger", command=self.clear_filters)  # add self.
        clear_btn.pack(side=tk.LEFT, padx=10)

        # Load PDF for Preview
        btn_load_pdf = ttk.Button(top_frame, text="📂 Load PDF", command=self.load_pdf)  # add self.
        btn_load_pdf.pack(side=tk.LEFT, padx=5)

        # PDF TO EXCEL Button
        pdf_to_excel_btn = tb.Button(top_frame, text="📥 PDF to Excel", bootstyle="info", command=self.convert_pdf_to_excel)  # add self.
        pdf_to_excel_btn.pack(side=tk.RIGHT, padx=10)

        # Create a Menu Button for Export Options
        export_menu_btn = tb.Menubutton(top_frame, text="📤 Export", bootstyle="warning")
        export_menu_btn.pack(side=tk.RIGHT, padx=10)

        # Create the Dropdown Menu
        export_menu = tk.Menu(export_menu_btn, tearoff=0)
        export_menu.add_command(label="📤 Export as CSV", command=lambda: self.export_filtered_data("csv"))  # add self.
        export_menu.add_command(label="📤 Export as Excel", command=lambda: self.export_filtered_data("xlsx"))  # add self.
        export_menu.add_command(label="📤 Export Full PDF", command=lambda: self.export_filtered_data("pdf"))  # add self.
        export_menu.add_command(label="📤 Export Individual PDFs", command=lambda: self.export_each_row_as_pdf())  # add self.

        # Attach the Menu to the Button
        export_menu_btn["menu"] = export_menu

        # 🟢 **Add Export Button**
        btn_export_pdf = tb.Button(top_frame, text="📤 Export PDFs", bootstyle="success",command=self.export_filled_pdfs)  # add self.
        btn_export_pdf.pack(side=tk.RIGHT, padx=10)

        # Treeview
        self.tree = ttk.Treeview(self)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def load_pdf(self):
        """Handles loading a PDF file."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            print(f"Loading PDF: {file_path}")
            # Add your PDF loading logic here
            try:
                # load the pdf.
                pdf_window = tk.Toplevel(self.parent)
                pdf_window.title("PDF Preview")
                PDFView(pdf_window, file_path).pack(fill=tk.BOTH, expand=True)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {e}")

    def convert_pdf_to_excel(self):
        """Handles converting a PDF to Excel."""
        print("Convert PDF to Excel called!")
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
                            # headers = first_row
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

    def export_filled_pdfs(self):
        """Handles exporting filled PDFs."""
        print("Export filled PDFs called!")
        if df is None:
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

        base_name = "Invoice"
        for index, row in df.iterrows():
            filled_pdf = fitz.open()

            for page in pdf_document:
                new_page = filled_pdf.new_page(width=page.rect.width, height=page.rect.height)
                new_page.show_pdf_page(new_page.rect, pdf_document, page.number)

                for box in box_data:
                    field = box["entry"]
                    field_column = box["column"].get()
                    if field_column in df.columns:
                        text_value = str(row[field_column])
                        x, y = pdf_canvas.coords(box["window"])
                        new_page.insert_text((x, y), text_value, fontsize=12, color=(0, 0, 0))

            output_file = os.path.join(save_folder, f"{base_name}_{index + 1}.pdf")
            filled_pdf.save(output_file)
            filled_pdf.close()

        messagebox.showinfo("Success", f"PDFs saved in {save_folder}")