import pandas as pd
from tkinter import ttk, messagebox

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
        messagebox.showinfo("No Results", "No matching records found.")
        return pd.DataFrame()

    return filtered_data

def display_data(tree, data, sort_orders):
    tree.delete(*tree.get_children())
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    for col in data.columns:
        arrow = ""
        if col in sort_orders and sort_orders[col] is not None:
            arrow = " ⬆" if sort_orders[col] else " ⬇"
        tree.heading(col, text=f"{col}{arrow}")
        tree.column(col, width=150, anchor="center")

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))