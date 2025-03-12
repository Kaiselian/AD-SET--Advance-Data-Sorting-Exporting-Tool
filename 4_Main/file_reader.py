import pandas as pd
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

def read_excel_csv(file_path, sheet_name=0, encoding="utf-8"):
    """
    Reads an Excel or CSV file and returns a DataFrame.

    Args:
        file_path (str): Path to the file.
        sheet_name (str/int): Name or index of the sheet to read (for Excel files).
        encoding (str): Encoding to use for reading CSV files.

    Returns:
        pd.DataFrame: DataFrame containing the file data, or None if an error occurs.
    """
    try:
        # Check if the file exists
        if not os.path.exists(file_path):
            logger.error(f"File not found at {file_path}")
            return None

        # Read CSV or Excel file
        if file_path.endswith(".csv"):
            logger.info(f"Reading CSV file: {file_path}")
            df = pd.read_csv(file_path, encoding=encoding, low_memory=False)
        elif file_path.endswith((".xlsx", ".xls")):
            logger.info(f"Reading Excel file: {file_path}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            logger.error("Unsupported file format. Please use Excel (.xlsx, .xls) or CSV (.csv).")
            return None

        # Validate the DataFrame
        if df.empty:
            logger.warning("The file is empty or no data was read.")
            return None

        logger.info(f"Successfully loaded {file_path}")
        return df

    except pd.errors.EmptyDataError:
        logger.error("The file is empty or no data was read.")
        return None
    except pd.errors.ParserError as e:
        logger.error(f"Error parsing the file: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error reading file: {e}")
        return None