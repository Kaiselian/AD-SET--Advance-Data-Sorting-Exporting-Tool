# data_filter.py
import pandas as pd
from tkinter import messagebox


def clear_filters(df, search_vars, sort_orders, display_data):
    """Reset filters and sorting, then refresh the displayed data."""
    if df is None:
        messagebox.showerror("Error", "No data loaded to clear filters.")
        return df, {}

    filtered_df = df.copy()
    search_vars["search_var"].set("")
    search_vars["sub_search_var"].set("")
    search_vars["column_var"].set("All Columns")
    search_vars["sub_search_column_var"].set("All Columns")
    search_vars["filter_var"].set("Contains")

    display_data(filtered_df)
    return filtered_df, {}


def apply_search_filter(df, search_query, column, filter_type):
    """Apply search filter based on query, column, and filter type."""
    if column == "All Columns":
        return df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)]

    if filter_type == "Contains":
        return df[df[column].astype(str).str.contains(search_query, case=False, na=False)]
    elif filter_type == "Equals":
        return df[df[column].astype(str) == search_query]
    elif filter_type == "Starts with":
        return df[df[column].astype(str).str.startswith(search_query, na=False)]


def search_and_generate(df, search_vars, display_data):
    """Perform search and filtering based on user input."""
    if df is None:
        messagebox.showerror("Error", "Please upload a file first.")
        return df

    main_query = search_vars["search_var"].get().strip()
    sub_query = search_vars["sub_search_var"].get().strip()
    main_column = search_vars["column_var"].get()
    sub_column = search_vars["sub_search_column_var"].get()
    filter_type = search_vars["filter_var"].get()

    if not main_query and not sub_query:
        messagebox.showerror("Error", "Please enter a search term.")
        return df

    filtered_data = df.copy()

    # Apply Main Search
    if main_query:
        filtered_data = apply_search_filter(filtered_data, main_query, main_column, filter_type)

    # Apply Sub-Search
    if sub_query:
        filtered_data = apply_search_filter(filtered_data, sub_query, sub_column, "Contains")

    display_data(filtered_data)
    return filtered_data
