import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox, ttk
import ttkbootstrap as tb
import os
import logging
import darkdetect
import sys
from datetime import datetime
from docx import Document
from file_reader import read_excel_csv
from data_mapper import scan_template_placeholders, prepare_row_data, replace_all_placeholders
from docx2pdf import convert

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


class DocumentFillerApp:
    def __init__(self, root):
        self.root = root
        self.load_default_templates()  # Load templates first
        self.setup_ui()
        self.setup_menu()

        # Initialize variables
        self.input_file = None
        self.output_folder = None
        self.current_data = None

    def load_default_templates(self):
        """Load default templates from the templates folder"""
        try:
            # Get the directory where the executable or script is located
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                application_path = os.path.dirname(sys.executable)
            else:
                # Running as script
                application_path = os.path.dirname(os.path.abspath(__file__))

            templates_dir = os.path.join(application_path, "templates")

            self.eligible_template = os.path.join(templates_dir, "eligible_template.docx")
            self.ineligible_template = os.path.join(templates_dir, "ineligible_template.docx")

            if not os.path.exists(self.eligible_template):
                raise FileNotFoundError(f"Eligible template not found at {self.eligible_template}")
            if not os.path.exists(self.ineligible_template):
                raise FileNotFoundError(f"Ineligible template not found at {self.ineligible_template}")

            logging.info("Default templates loaded successfully")

        except Exception as e:
            logging.error(f"Failed to load default templates: {str(e)}")
            messagebox.showerror("Error", f"Failed to load default templates: {str(e)}")
            self.root.destroy()

    def setup_ui(self):
        """Setup the main user interface"""
        self.root.title("Automated ISD Document Generator")
        self.root.geometry("1200x800")

        # Main container
        main_frame = tb.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Left panel - Controls
        control_frame = tb.Frame(main_frame)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # Control buttons
        btn_data = tb.Button(control_frame, text="📂 Upload Data File", command=self.upload_data_file)
        btn_data.pack(fill=tk.X, padx=10, pady=5)

        btn_output = tb.Button(control_frame, text="📁 Select Output Folder", command=self.select_output_folder)
        btn_output.pack(fill=tk.X, padx=10, pady=5)

        btn_start = tb.Button(control_frame, text="🚀 Generate ISD Invoices", bootstyle="success",
                              command=self.start_processing)
        btn_start.pack(fill=tk.X, padx=10, pady=20)

        # Template status labels
        self.lbl_eligible_template = tb.Label(control_frame,
                                              text=f"✅ Eligible Template: {os.path.basename(self.eligible_template)}",
                                              bootstyle="success")
        self.lbl_eligible_template.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_ineligible_template = tb.Label(control_frame,
                                                text=f"✅ Ineligible Template: {os.path.basename(self.ineligible_template)}",
                                                bootstyle="success")
        self.lbl_ineligible_template.pack(fill=tk.X, padx=10, pady=5)

        # Status labels
        self.lbl_data = tb.Label(control_frame, text="No Data File Loaded", bootstyle="secondary")
        self.lbl_data.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_output = tb.Label(control_frame, text="No Output Folder Selected", bootstyle="secondary")
        self.lbl_output.pack(fill=tk.X, padx=10, pady=5)

        # Right panel - Data Preview
        preview_frame = tb.Frame(main_frame)
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        preview_label = tb.Label(preview_frame, text="Data Preview", bootstyle="primary")
        preview_label.pack(fill=tk.X, pady=5)

        # Create the treeview
        self.tree = self.create_treeview(preview_frame)

    def start_processing(self):
        """Start the document generation process"""
        if not all([self.input_file, self.output_folder]):
            messagebox.showerror("Error", "Please select data file and output folder!")
            return

        try:
            data = read_excel_csv(self.input_file)
            if data is None:
                messagebox.showerror("Error", "Failed to read data file.")
                return

            # Create output folders
            pdf_output_folder = os.path.join(self.output_folder, "PDF_Output")
            os.makedirs(pdf_output_folder, exist_ok=True)

            temp_docx_folder = os.path.join(self.output_folder, "TEMP_DOCX")
            os.makedirs(temp_docx_folder, exist_ok=True)

            success_count = 0

            for idx, row in data.iterrows():
                try:
                    # Determine template type
                    is_eligible = self.is_row_eligible(row)
                    template_path = self.eligible_template if is_eligible else self.ineligible_template
                    prefix = "Eligible" if is_eligible else "Ineligible"

                    # Generate document
                    doc = Document(template_path)
                    placeholders = scan_template_placeholders(template_path)
                    row_data = prepare_row_data(row, placeholders, is_eligible)

                    if not replace_all_placeholders(doc, row_data):
                        logging.error(f"Skipping row {idx} due to replacement errors")
                        continue

                    # Save temporary DOCX
                    invoice_num = str(row.get('INVOICE_NUMBER', idx + 1)).strip()
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    docx_filename = f"{prefix}_ISD_{invoice_num}_{timestamp}.docx"
                    docx_path = os.path.join(temp_docx_folder, docx_filename)
                    doc.save(docx_path)

                    # Convert to PDF
                    pdf_filename = f"{prefix}_ISD_{invoice_num}_{timestamp}.pdf"
                    pdf_path = os.path.join(pdf_output_folder, pdf_filename)
                    convert(docx_path, pdf_path)

                    # Delete temporary DOCX
                    os.remove(docx_path)

                    success_count += 1
                    logging.info(f"Generated {pdf_filename}")

                except Exception as e:
                    logging.error(f"Error processing row {idx}: {str(e)}", exc_info=True)
                    continue

            # Remove temporary DOCX folder if empty
            try:
                os.rmdir(temp_docx_folder)
            except OSError:
                pass  # Folder not empty

            messagebox.showinfo("Success",
                                f"Processing complete!\n\nGenerated {success_count} PDF invoices.\n"
                                f"Location: {pdf_output_folder}")

        except Exception as e:
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
            logging.error(f"Processing error: {str(e)}")

    def is_row_eligible(self, row):
        """Determine if row contains eligible or ineligible data"""
        eligible_cols = [
            'ELIGIBLE_CGST', 'ELIGIBLE_SGST', 'ELIGIBLE_UTGST', 'ELIGIBLE_IGST',
            'ELIGIBLECGST', 'ELIGIBLESGST', 'ELIGIBLEUTGST', 'ELIGIBLEIGST'
        ]

        # Check if any eligible tax amount is > 0
        for col in eligible_cols:
            if col in row:
                try:
                    val = float(row[col]) if pd.notna(row[col]) else 0
                    if val > 0:
                        return True
                except (ValueError, TypeError):
                    continue
        return False

    def create_treeview(self, parent_frame):
        """Create and configure the Treeview widget"""
        tree_frame = tb.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Scrollbars
        tree_scroll_y = tk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = tk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview
        tree = ttk.Treeview(tree_frame, style="Custom.Treeview",
                            yscrollcommand=tree_scroll_y.set,
                            xscrollcommand=tree_scroll_x.set)
        tree.pack(pady=10, fill=tk.BOTH, expand=True)

        # Configure scrollbars
        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)

        return tree

    def display_data(self, data):
        """Display data in the Treeview"""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(data.columns)
        self.tree["show"] = "headings"

        for col in data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        for _, row in data.iterrows():
            self.tree.insert("", "end", values=list(row))

        self.tree.update_idletasks()

    def setup_menu(self):
        """Setup the menu bar"""
        menu_bar = tk.Menu(self.root)

        # File menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Upload Data File", command=self.upload_data_file)
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Theme menu
        theme_menu = tk.Menu(menu_bar, tearoff=0)
        theme_options = {
            "darkly": "🌙 Dark",
            "journal": "📖 Light",
            "flatly": "📄 Flat",
            "cyborg": "🤖 Cyborg",
            "superhero": "🦸 Superhero",
            "minty": "🌿 Minty"
        }

        for theme, label in theme_options.items():
            theme_menu.add_command(label=label, command=lambda t=theme: self.change_theme(t))

        menu_bar.add_cascade(label="Theme", menu=theme_menu)

        self.root.config(menu=menu_bar)

    def change_theme(self, selected_theme):
        """Change the application theme"""
        self.root.style.theme_use(selected_theme)

    def upload_data_file(self):
        """Handle data file upload"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.xls;*.csv")])
        if file_path:
            self.input_file = file_path
            self.lbl_data.config(text=f"📂 {os.path.basename(file_path)} Loaded")
            logging.info(f"Data file loaded: {file_path}")

            try:
                self.current_data = read_excel_csv(file_path)
                if self.current_data is not None:
                    self.display_data(self.current_data)
                    messagebox.showinfo("Success", "Data file loaded and displayed successfully!")
                else:
                    messagebox.showerror("Error", "Failed to read data file.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load data: {str(e)}")
                logging.error(f"Data loading error: {str(e)}")

    def select_output_folder(self):
        """Handle output folder selection"""
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.lbl_output.config(text=f"📁 Output Folder: {folder}")
            logging.info(f"Output folder selected: {folder}")


# Initialize and run the application
if __name__ == "__main__":
    theme = "darkly" if darkdetect.isDark() else "journal"
    root = tb.Window(themename=theme)
    app = DocumentFillerApp(root)
    root.mainloop()