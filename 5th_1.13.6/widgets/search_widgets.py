import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb

class SearchWidgets:
    def __init__(self, parent, tree, display_data_callback):
        self.parent = parent
        self.tree = tree
        self.display_data_callback = display_data_callback
        self.search_var = tk.StringVar()
        self.sub_search_var = tk.StringVar()
        self.column_var = tk.StringVar(value="All Columns")
        self.sub_search_column_var = tk.StringVar(value="All Columns")
        self.filter_var = tk.StringVar(value="Contains")
        self.create_widgets()

    def create_widgets(self):
        """Creates the search widgets."""

        # Main Search Bar
        self.search_entry = tb.Entry(self.parent, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=10)

        # Sub-Search Bar
        self.sub_search_entry = tb.Entry(self.parent, textvariable=self.sub_search_var, width=40)
        self.sub_search_entry.pack(side=tk.LEFT, padx=10)

        # Column Dropdowns
        self.column_dropdown = ttk.Combobox(self.parent, textvariable=self.column_var, state="readonly")
        self.column_dropdown.pack(side=tk.LEFT, padx=10)

        self.sub_search_column_dropdown = ttk.Combobox(self.parent, textvariable=self.sub_search_column_var, state="readonly")
        self.sub_search_column_dropdown.pack(side=tk.LEFT, padx=10)

        # Filter Type Dropdown
        self.filter_dropdown = ttk.Combobox(self.parent, textvariable=self.filter_var, state="readonly", values=["Contains", "Equals", "Starts with"])
        self.filter_dropdown.pack(side=tk.LEFT, padx=10)

        # Search Button
        self.search_btn = tb.Button(self.parent, text="🔍 Search", bootstyle="success", command=self.perform_search)
        self.search_btn.pack(side=tk.LEFT, padx=10)

        # Clear Filters Button
        self.clear_btn = tb.Button(self.parent, text="Clear Filters", bootstyle="danger", command=self.clear_filters)
        self.clear_btn.pack(side=tk.LEFT, padx=10)

    # Clear Data Searched and Reset Sorting
    def clear_filters(self):
        """Resets all search filters and refreshes the dataset."""
        self.search_var.set("")
        self.sub_search_var.set("")
        self.column_var.set("All Columns")
        self.sub_search_column_var.set("All Columns")
        self.filter_var.set("Contains")

        self.display_data_callback(
            search_query=self.search_var.get(),
            sub_query=self.sub_search_var.get(),
            main_column=self.column_var.get(),
            sub_column=self.sub_search_column_var.get(),
            filter_type=self.filter_var.get()
        )

    # 🔍 Combined Search (Main & Sub-Search)
    def search_and_generate(self):
        """Filters data based on search criteria and updates the Treeview."""
        search_query = self.search_var.get().strip()
        sub_query = self.sub_search_var.get().strip()
        main_column = self.column_var.get()
        sub_column = self.sub_search_column_var.get()
        filter_type = self.filter_var.get()

        self.update_display_callback(
            search_query=search_query,
            sub_query=sub_query,
            main_column=main_column,
            sub_column=sub_column,
            filter_type=filter_type
        )

        # 🔹 Apply Main Search
        if main_query:
            if main_column == "All Columns":
                filtered_data = filtered_data[
                    filtered_data.apply(
                        lambda row: row.astype(str).str.contains(main_query, case=False, na=False).any(), axis=1)
                ]
            else:
                if filter_type == "Contains":
                    filtered_data = filtered_data[
                        filtered_data[main_column].astype(str).str.contains(main_query, case=False, na=False)]
                elif filter_type == "Equals":
                    filtered_data = filtered_data[filtered_data[main_column].astype(str) == main_query]
                elif filter_type == "Starts with":
                    filtered_data = filtered_data[
                        filtered_data[main_column].astype(str).str.startswith(main_query, na=False)]

        # 🔹 Apply Sub-Search on Filtered Data
        if sub_query:
            if sub_column == "All Columns":
                filtered_data = filtered_data[
                    filtered_data.apply(lambda row: row.astype(str).str.contains(sub_query, case=False, na=False).any(),
                                        axis=1)
                ]
            else:
                filtered_data = filtered_data[
                    filtered_data[sub_column].astype(str).str.contains(sub_query, case=False, na=False)]

        # 🛑 FIXED: Correct placement of "No Results" message
        if filtered_data.empty:
            messagebox.showinfo("No Results", "No matching records found.")
            return

        display_data(filtered_data)  # ✅ Display only once
        filtered_df = filtered_data  # ✅ Store for sorting

        # Store filtered data for sorting

        if filtered_data.empty:
            messagebox.showinfo("No Results", "No matching records found.")
            return

        display_data(filtered_data)

    def perform_search(self):
        """Triggers the search and updates the Treeview."""
        search_query = self.search_var.get().strip()
        sub_query = self.sub_search_var.get().strip()
        main_column = self.column_var.get()
        sub_column = self.sub_search_column_var.get()
        filter_type = self.filter_var.get()

        self.display_data_callback(search_query, sub_query, main_column, sub_column, filter_type)