from PyQt5.QtWidgets import QWidget, QPushButton, QMenu, QHBoxLayout

class ExportWidgets(QWidget):
    def __init__(self, parent, filtered_df, export_data_callback):
        super().__init__(parent)
        self.filtered_df = filtered_df
        self.export_data_callback = export_data_callback
        self.create_widgets()

    def create_widgets(self):
        layout = QHBoxLayout(self)

        export_menu_btn = QPushButton("📤 Export")
        export_menu = QMenu(self)
        export_menu.addAction("📤 Export as CSV", lambda: self.export_data_callback("csv"))
        export_menu.addAction("📤 Export as Excel", lambda: self.export_data_callback("xlsx"))
        export_menu.addAction("📤 Export as PDF", lambda: self.export_data_callback("pdf"))
        export_menu_btn.setMenu(export_menu)

        layout.addWidget(export_menu_btn)
        self.setLayout(layout)