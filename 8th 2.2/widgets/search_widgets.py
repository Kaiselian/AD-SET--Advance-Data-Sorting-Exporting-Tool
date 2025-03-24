from PyQt5.QtWidgets import QWidget, QLineEdit, QComboBox, QPushButton, QHBoxLayout
from PyQt5.QtCore import Qt

class SearchWidgets(QWidget):
    def __init__(self, parent, tree, display_data_callback):
        super().__init__(parent)
        self.tree = tree
        self.display_data_callback = display_data_callback
        self.create_widgets()

    def create_widgets(self):
        layout = QHBoxLayout(self)

        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Search...")
        layout.addWidget(self.search_entry)

        self.sub_search_entry = QLineEdit()
        self.sub_search_entry.setPlaceholderText("Sub-Search...")
        layout.addWidget(self.sub_search_entry)

        self.column_dropdown = QComboBox()
        self.column_dropdown.addItem("All Columns")
        layout.addWidget(self.column_dropdown)

        self.sub_search_column_dropdown = QComboBox()
        self.sub_search_column_dropdown.addItem("All Columns")
        layout.addWidget(self.sub_search_column_dropdown)

        self.filter_dropdown = QComboBox()
        self.filter_dropdown.addItems(["Contains", "Equals", "Starts with"])
        layout.addWidget(self.filter_dropdown)

        self.search_btn = QPushButton("🔍Search")
        self.search_btn.clicked.connect(self.perform_search)
        layout.addWidget(self.search_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_filters)
        layout.addWidget(self.clear_btn)

        self.setLayout(layout)

    def perform_search(self):
        search_query = self.search_entry.text().strip()
        sub_query = self.sub_search_entry.text().strip()
        main_column = self.column_dropdown.currentText()
        sub_column = self.sub_search_column_dropdown.currentText()
        filter_type = self.filter_dropdown.currentText()
        self.display_data_callback(search_query, sub_query, main_column, sub_column, filter_type)

    def clear_filters(self):
        self.search_entry.setText("")
        self.sub_search_entry.setText("")
        self.column_dropdown.setCurrentText("All Columns")
        self.sub_search_column_dropdown.setCurrentText("All Columns")
        self.filter_dropdown.setCurrentText("Contains")
        self.display_data_callback(
            search_query="",
            sub_query="",
            main_column="All Columns",
            sub_column="All Columns",
            filter_type="Contains"
        )

    def update_columns(self, columns):
        self.column_dropdown.clear()
        self.column_dropdown.addItem("All Columns")
        self.column_dropdown.addItems(columns)
        self.sub_search_column_dropdown.clear()
        self.sub_search_column_dropdown.addItem("All Columns")
        self.sub_search_column_dropdown.addItems(columns)