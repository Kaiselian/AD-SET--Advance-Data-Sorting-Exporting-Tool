from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton
from utils.invoice_generator import InvoiceGenerator

class InvoiceView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.invoice_generator = InvoiceGenerator()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Upload Button
        self.upload_btn = QPushButton("📂 Upload Invoice Data")
        self.upload_btn.clicked.connect(self.upload_file)
        layout.addWidget(self.upload_btn)

        # Generate Invoice Button
        self.generate_btn = QPushButton("📄 Generate Invoices")
        self.generate_btn.clicked.connect(self.generate_invoices)
        layout.addWidget(self.generate_btn)

        self.setLayout(layout)

    def upload_file(self):
        """Load invoice data"""
        self.invoice_generator.load_invoice_data()

    def generate_invoices(self):
        """Generate invoices"""
        self.invoice_generator.generate_invoices()