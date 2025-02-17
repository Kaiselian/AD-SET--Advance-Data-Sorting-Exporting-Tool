import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import ttkbootstrap as tb  # Modern UI Framework

# Global variables
df = None

# Global variable to store the last filtered dataset
filtered_df = None

def sort_column(order):
    global filtered_df
    selected_column = column_var.get()

    # Ensure there's data to sort and a valid column is selected
    if filtered_df is not None and selected_column in filtered_df.columns:
        filtered_df = filtered_df.sort_values(by=selected_column, ascending=(order == "asc"))
        display_data(filtered_df)  # Display only the sorted filtered data


# Initialize GUI
root = tb.Window(themename="darkly")  # Default theme, fixed
root.title("Advanced Data Search & Export Tool 1.06")
root.geometry("1920x1080")
root.state("zoomed")


# 🟢 Upload File Function
def upload_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False) if file_path.endswith(".csv") else pd.read_excel(file_path, sheet_name=0)

            if df.empty:
                messagebox.showerror("Error", "Loaded file is empty or could not be read.")
                return

            update_columns()
            display_data(df)  # Now tree exists, so no error.
            messagebox.showinfo("Success", "File uploaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            print("Upload Error:", e)  # Debugging

# 🔄 Display Data in Treeview with Sorting Buttons
def display_data(data):
    global filtered_df
    filtered_df = data  # Store filtered data

    tree.delete(*tree.get_children())  # Clear existing data
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    # Loop through columns and add sorting icons only to the last sorted column
    for col in data.columns:
        arrow = ""
        if col in sort_orders and sort_orders[col] is not None:  # Show arrow only if column was sorted
            arrow = " ⬆" if sort_orders[col] else " ⬇"
        tree.heading(col, text=f"{col}{arrow}", command=lambda c=col: toggle_sort_order(c))
        tree.column(col, width=150, anchor="center")

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

# 🔼🔽 Toggle Sort Order
sort_orders = {}  # Dictionary to track column sorting order

def toggle_sort_order(column):
    global filtered_df

    if filtered_df is None or column not in filtered_df.columns:
        messagebox.showerror("Error", "Invalid column selection.")
        return

    # Reset all sorting indicators except the selected column
    for col in sort_orders:
        if col != column:
            sort_orders[col] = None  # Remove sorting order for other columns

    # Toggle sorting order for the selected column
    sort_orders[column] = not sort_orders.get(column, False)
    ascending = sort_orders[column]

    # Sort the data based on the selected column and order
    filtered_df = filtered_df.sort_values(by=column, ascending=ascending)

    display_data(filtered_df)  # Refresh sorted data


# Clear Data Searched and Reset Sorting
def clear_filters():
    global filtered_df, sort_orders
    if df is None:
        messagebox.showerror("Error", "No data loaded to clear filters.")
        return

    filtered_df = df.copy()  # Reset data
    sort_orders = {}  # Reset sorting order

    search_var.set("")
    sub_search_var.set("")
    column_var.set("All Columns")
    sub_search_column_var.set("All Columns")
    filter_var.set("Contains")

    display_data(filtered_df)  # Refresh without sorting icons


# 🔍 Combined Search (Main & Sub-Search)
def search_and_generate():
    global df, filtered_df

    if df is None:
        messagebox.showerror("Error", "Please upload a file first.")
        return

    main_query = search_var.get().strip()
    sub_query = sub_search_var.get().strip()
    main_column = column_var.get()
    sub_column = sub_search_column_var.get()
    filter_type = filter_var.get()

    if not main_query and not sub_query:
        messagebox.showerror("Error", "Please enter a search term.")
        return

    filtered_data = df.copy()

    # 🔹 Apply Main Search
    if main_query:
        if main_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(main_query, case=False, na=False).any(), axis=1)
            ]
        else:
            if filter_type == "Contains":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.contains(main_query, case=False, na=False)]
            elif filter_type == "Equals":
                filtered_data = filtered_data[filtered_data[main_column].astype(str) == main_query]
            elif filter_type == "Starts with":
                filtered_data = filtered_data[filtered_data[main_column].astype(str).str.startswith(main_query, na=False)]

    # 🔹 Apply Sub-Search on Filtered Data
    if sub_query:
        if sub_column == "All Columns":
            filtered_data = filtered_data[
                filtered_data.apply(lambda row: row.astype(str).str.contains(sub_query, case=False, na=False).any(), axis=1)
            ]
        else:
            filtered_data = filtered_data[filtered_data[sub_column].astype(str).str.contains(sub_query, case=False, na=False)]

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

# 📤 Export Data
def export_filtered_data(format):
    if filtered_df is None:
        messagebox.showerror("Error", "No filtered data to export.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=f".{format}", filetypes=[(f"{format.upper()} files", f"*.{format}")])
    if save_path:
        try:
            if format == "xlsx":
                filtered_df.to_excel(save_path, index=False)
            elif format == "csv":
                filtered_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Filtered data saved as {format.upper()} successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")


# 🔹 UI Layout - Top Bar
top_frame = tb.Frame(root)
top_frame.pack(pady=10, fill=tk.X, padx=20)

upload_btn = tb.Button(top_frame, text="📂 Upload File", bootstyle="primary", command=upload_file)
upload_btn.pack(side=tk.LEFT, padx=10)

search_var = tk.StringVar()
search_entry = tb.Entry(top_frame, textvariable=search_var, width=40)
search_entry.pack(side=tk.LEFT, padx=10)
search_entry.bind("<Return>", lambda event: search_and_generate())  # ENTER triggers search

search_btn = tb.Button(top_frame, text="🔍", bootstyle="success", command=search_and_generate)
search_btn.pack(side=tk.LEFT, padx=10)

# Clear Button
clear_btn = tb.Button(top_frame, text="❌ Clear Filters", bootstyle="danger", command=clear_filters)
clear_btn.pack(side=tk.LEFT, padx=10)

# 📤 Export Buttons
export_filtered_csv_btn = tb.Button(top_frame, text="📤 Export Filtered CSV", bootstyle="warning", command=lambda: export_filtered_data("csv"))
export_filtered_csv_btn.pack(side=tk.RIGHT, padx=10)

export_filtered_xlsx_btn = tb.Button(top_frame, text="📤 Export Filtered Excel", bootstyle="warning", command=lambda: export_filtered_data("xlsx"))
export_filtered_xlsx_btn.pack(side=tk.RIGHT, padx=10)

# 🔍 Sub-Search Bar & Column Selection
sub_search_var = tk.StringVar()
sub_search_entry = tb.Entry(top_frame, textvariable=sub_search_var, width=40)
sub_search_entry.pack(side=tk.LEFT, padx=10)
sub_search_entry.bind("<Return>", lambda event: search_and_generate())  # ENTER triggers sub-search

sub_search_column_var = tk.StringVar(value="All Columns")
sub_search_column_dropdown = ttk.Combobox(top_frame, textvariable=sub_search_column_var, state="readonly")
sub_search_column_dropdown.pack(side=tk.LEFT, padx=10)

sub_search_btn = tb.Button(top_frame, text="🔍 Sub-Search", bootstyle="success", command=search_and_generate)
sub_search_btn.pack(side=tk.LEFT, padx=10)

# 🔽 Column Dropdown
column_var = tk.StringVar(value="All Columns")
column_dropdown = ttk.Combobox(top_frame, textvariable=column_var, state="readonly")
# column_dropdown.pack(side=tk.LEFT, padx=10)


# 🔍 Filter Type Dropdown
filter_var = tk.StringVar(value="Contains")
filter_dropdown = ttk.Combobox(top_frame, textvariable=filter_var, state="readonly", values=["Contains", "Equals", "Starts with"])
# filter_dropdown.pack(side=tk.LEFT, padx=10)

# 🔹 Treeview for Data Display
frame2 = tb.Frame(root)
frame2.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

tree = ttk.Treeview(frame2, style="Custom.Treeview")
tree.pack(pady=10, fill=tk.BOTH, expand=True)

# 🔄 Update column dropdown when a file is loaded
def update_columns():
    if df is not None:
        column_dropdown["values"] = ["All Columns"] + list(df.columns)
        column_var.set("All Columns")
        sub_search_column_dropdown["values"] = ["All Columns"] + list(df.columns)
        sub_search_column_var.set("All Columns")

root.mainloop()
