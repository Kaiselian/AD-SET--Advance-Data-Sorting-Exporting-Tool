import pandas as pd

def display_data(tree, data, sort_orders=None):
    """Displays DataFrame data in the Treeview table."""
    tree.delete(*tree.get_children())  # Clear existing data
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    # Add column headers
    for col in data.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    # Insert data into the table
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

    tree.update_idletasks()  # Refresh UI
def filter_data(df, search_query, sub_query, main_column, sub_column, filter_type):
    """Filters the DataFrame based on search criteria."""
    filtered_df = df.copy()

    # 🔹 Apply Main Search
    if search_query:
        if main_column == "All Columns":
            filtered_df = filtered_df[
                filtered_df.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
            ]
        else:
            if filter_type == "Contains":
                filtered_df = filtered_df[filtered_df[main_column].astype(str).str.contains(search_query, case=False, na=False)]
            elif filter_type == "Equals":
                filtered_df = filtered_df[filtered_df[main_column].astype(str) == search_query]
            elif filter_type == "Starts with":
                filtered_df = filtered_df[filtered_df[main_column].astype(str).str.startswith(search_query, na=False)]

    # 🔹 Apply Sub-Search
    if sub_query:
        if sub_column == "All Columns":
            filtered_df = filtered_df[
                filtered_df.apply(lambda row: row.astype(str).str.contains(sub_query, case=False, na=False).any(), axis=1)
            ]
        else:
            filtered_df = filtered_df[filtered_df[sub_column].astype(str).str.contains(sub_query, case=False, na=False)]

    return filtered_df