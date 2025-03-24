from PyQt5.QtWidgets import QWidget, QLabel, QScrollArea, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLineEdit
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from utils.pdf_utils import load_pdf

class PDFView(QWidget):
    def __init__(self, parent, pdf_path):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.text_boxes = []
        self.box_data = []
        self.pdf_pixmap = None
        self.pdf_document = None
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout(self)

        self.pdf_pixmap, self.pdf_document = load_pdf(self.pdf_path)

        if self.pdf_pixmap:
            self.scroll_area = QScrollArea()
            self.scroll_area.setAlignment(Qt.AlignCenter)
            self.scroll_area.setWidgetResizable(True)

            self.pdf_label = QLabel()
            self.pdf_label.setPixmap(self.pdf_pixmap)
            self.scroll_area.setWidget(self.pdf_label)

            layout.addWidget(self.scroll_area, 4)

            self.frame_right = QWidget()
            self.frame_right.setStyleSheet("background-color: #f0f0f0;")
            right_layout = QVBoxLayout(self.frame_right)

            self.btn_add_box = QPushButton("➕ Add Text Box")
            self.btn_add_box.clicked.connect(self.add_text_box)
            right_layout.addWidget(self.btn_add_box)

            layout.addWidget(self.frame_right, 1)
        else:
            error_label = QLabel("Failed to load PDF.")
            layout.addWidget(error_label)

        self.setLayout(layout)

    def add_text_box(self):
        frame = QWidget()
        frame_layout = QHBoxLayout(frame)

        entry = QLineEdit()
        frame_layout.addWidget(entry)

        self.text_boxes.append(entry)

        # Create a placeholder label for positioning
        placeholder_label = QLabel()
        placeholder_label.setText("Placeholder")
        placeholder_label.setStyleSheet("border: 1px solid red;")
        placeholder_label.setFixedSize(100,20)

        # Add to the scroll area
        self.pdf_label.layout = QVBoxLayout()
        self.pdf_label.setLayout(self.pdf_label.layout)
        self.pdf_label.layout.addWidget(placeholder_label)

        # Store the data
        self.box_data.append({"entry": entry, "label": placeholder_label, "x": 50, "y": 50, "column": None})