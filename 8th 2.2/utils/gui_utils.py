from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5.QtCore import Qt

def create_table_widget(parent):
    """
    Creates a QTableWidget with scrollbars.
    """
    table = QTableWidget(parent)
    table.setAlternatingRowColors(True)
    table.setEditTriggers(QTableWidget.NoEditTriggers)  # Make table read-only
    table.setSelectionBehavior(QTableWidget.SelectRows)  # Select entire rows
    table.setSelectionMode(QTableWidget.SingleSelection)  # Single row selection

    # Configure horizontal header
    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.ResizeToContents)  # Resize columns to content

    return table

def display_data(table, data, sort_orders=None):
    """
    Displays data in the QTableWidget.
    """
    if sort_orders is None:
        sort_orders = {}

    table.setRowCount(0)  # Clear existing data
    if data.empty:
        table.setColumnCount(0)
        return

    table.setColumnCount(len(data.columns))
    table.setHorizontalHeaderLabels(list(data.columns))

    for i, row in data.iterrows():
        table.insertRow(i)
        for j, col in enumerate(data.columns):
            item = QTableWidgetItem(str(row[col]))
            item.setTextAlignment(Qt.AlignCenter)  # Align text to center
            table.setItem(i, j, item)

    # Adjust column widths to content

    table.resizeColumnsToContents()