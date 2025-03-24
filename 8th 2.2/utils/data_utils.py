# (Your original data_utils.py content)
import pandas as pd
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem

def filter_data(df, search_query, sub_query, main_column, sub_column, filter_type):
    filtered_data = df.copy()

    if search_query:
        if main_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
            ]
        else:
            if filter_type == "Contains":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.contains(search_query, case=False, na=False)]
            elif filter_type == "Equals":
                filtered_data = filtered_data[filtered_data[main_column].astype(str) == search_query]
            elif filter_type == "Starts with":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.startswith(search_query, na=False)]

    if sub_query:
        if sub_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(sub_query, case=False, na=False).any(), axis=1)
            ]
        else:
            filtered_data = filtered_data[filtered_data[sub_column].astype(str).str.contains(sub_query, case=False, na=False)]

    if filtered_data.empty:
        QMessageBox.showinfo("No Results", "No matching records found.")
        return pd.DataFrame()

    return filtered_data

def display_data(table, data, sort_orders):
    table.setRowCount(0)
    table.setColumnCount(len(data.columns))
    table.setHorizontalHeaderLabels(list(data.columns))

    for i, row in data.iterrows():
        for j, col in enumerate(data.columns):
            item = QTableWidgetItem(str(row[col]))
            table.setItem(i, j, item)