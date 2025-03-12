import pandas as pd


def read_excel_csv(file_path):
    """
    Reads an Excel or CSV file and returns a DataFrame.

    :param file_path: Path to the data file
    :return: Pandas DataFrame or None if an error occurs
    """
    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False)
        elif file_path.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path, sheet_name=0)
        else:
            print("Unsupported file format. Please use Excel or CSV.")
            return None

        if df.empty:
            print("Error: The data file is empty.")
            return None

        print(f"✅ Successfully loaded {file_path}")
        return df

    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None
