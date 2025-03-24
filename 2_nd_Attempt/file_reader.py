import pandas as pd
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def read_excel_csv(file_path, sheet_name=None):
    """
    Reads an Excel or CSV file and returns a DataFrame.

    :param file_path: Path to the data file
    :param sheet_name: Name of the sheet to read (for Excel files). If None, reads the first sheet.
    :return: Pandas DataFrame or None if an error occurs
    """
    try:
        if file_path.endswith(".csv"):
            logging.info(f"Reading CSV file: {file_path}")
            df = pd.read_csv(file_path, encoding="utf-8", low_memory=False)
        elif file_path.endswith((".xlsx", ".xls")):
            logging.info(f"Reading Excel file: {file_path}")
            df = pd.read_excel(file_path, sheet_name=sheet_name if sheet_name else 0)
        else:
            logging.error("Unsupported file format. Please use Excel (.xlsx, .xls) or CSV (.csv).")
            return None

        if df.empty:
            logging.error("Error: The data file is empty.")
            return None

        logging.info(f"✅ Successfully loaded {file_path}")
        return df

    except FileNotFoundError:
        logging.error(f"❌ Error: File not found at {file_path}")
        return None
    except pd.errors.EmptyDataError:
        logging.error("❌ Error: The data file is empty or corrupted.")
        return None
    except Exception as e:
        logging.error(f"❌ Error reading file: {e}")
        return None

# Example usage
if __name__ == "__main__":
    file_path = "C:/Users/anich/Downloads/Fields.xlsx"  # Replace with your file path
    df = read_excel_csv(file_path)
    if df is not None:
        print(df.head())  # Display the first few rows of the DataFrame