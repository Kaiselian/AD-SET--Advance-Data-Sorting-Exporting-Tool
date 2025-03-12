import pandas as pd
import os

def read_excel_csv(file_path, sheet_name=0):
    """
    Reads an Excel or CSV file and returns a DataFrame.

    :param file_path: Path to the data file
    :param sheet_name: sheet name or index (for Excel). Defaults to 0.
    :return: Pandas DataFrame or None if an error occurs
    """
    try:
        if not os.path.exists(file_path):
            print(f"❌ Error: File not found at {file_path}")
            return None

        # Read CSV file with automatic encoding detection
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, encoding="utf-8-sig", low_memory=False, dtype=str)

        # Read Excel file
        elif file_path.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)

        else:
            print("❌ Unsupported file format. Please use Excel (.xls, .xlsx) or CSV.")
            return None

        if df.empty:
            print("❌ Error: The data file is empty.")
            return None

        print(f"✅ Successfully loaded {file_path} with {len(df)} rows and {len(df.columns)} columns.")
        return df

    except FileNotFoundError:
        print(f"❌ Error: File not found at {file_path}")
        return None
    except pd.errors.ParserError:
        print(f"❌ Error: Could not parse CSV file. Check for formatting errors in {file_path}")
        return None
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None
