import os
import pandas as pd
import logging
from typing import Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def read_excel_csv(file_path: str) -> Optional[pd.DataFrame]:
    """
    Reads Excel or CSV file into a pandas DataFrame with robust error handling.

    Args:
        file_path: Path to the input file (Excel or CSV)

    Returns:
        pandas DataFrame if successful, None otherwise
    """
    try:
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            return None

        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext in ['.xlsx', '.xls']:
            # Read Excel with formula evaluation and proper dtype handling
            df = pd.read_excel(
                file_path,
                engine='openpyxl',
                dtype=str,  # Read all as string to preserve formatting
                na_values=['', 'NA', 'N/A', 'NULL'],
                keep_default_na=False
            )
            logging.info(f"Successfully loaded Excel file: {file_path}")

        elif file_ext == '.csv':
            # Read CSV with flexible parsing
            df = pd.read_csv(
                file_path,
                dtype=str,
                encoding='utf-8',
                na_values=['', 'NA', 'N/A', 'NULL'],
                keep_default_na=False
            )
            logging.info(f"Successfully loaded CSV file: {file_path}")

        else:
            logging.error(f"Unsupported file format: {file_path}")
            return None

        # Clean column names and data
        df = clean_data(df)
        logging.info(f"Columns in data: {df.columns.tolist()}")
        logging.info(f"First row sample:\n{df.iloc[0].to_dict()}")

        return df

    except PermissionError:
        logging.error(f"Permission denied when reading: {file_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading {file_path}: {str(e)}")
        return None


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans and normalizes the loaded DataFrame.

    Args:
        df: Raw pandas DataFrame

    Returns:
        Cleaned DataFrame with normalized column names and values
    """
    # Normalize column names
    df.columns = [
        col.strip()
        .upper()
        .replace(' ', '_')
        .replace('-', '_')
        .replace('.', '')
        for col in df.columns
    ]

    # Clean string values
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
            df[col] = df[col].replace({'': None, 'nan': None, 'None': None})

    # Convert amount columns to numeric
    amount_cols = ['AMOUNT', 'CGST', 'SGST', 'UTGST', 'IGST']
    for col in amount_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df