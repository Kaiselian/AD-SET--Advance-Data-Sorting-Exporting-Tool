import os
import pandas as pd
import logging
from typing import Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def read_excel_csv(file_path: str) -> Optional[pd.DataFrame]:
    try:
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            return None

        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext in ['.xlsx', '.xls']:
            # Read the second row to get tax type labels
            header_df = pd.read_excel(file_path, header=None, nrows=2)
            tax_labels = header_df.iloc[1, 14:22].tolist()  # Columns O to V

            # Read data with proper column names
            df = pd.read_excel(
                file_path,
                engine='openpyxl',
                header=0,
                skiprows=[1],  # Skip the tax type labels row
                names=[
                    'INVOICE_NUMBER', 'INVOICE_DATE', 'ISD_DISTRIBUTOR_GSTIN',
                    'ISD_DISTRIBUTOR_NAME', 'ISD_DISTRIBUTOR_ADDRESS',
                    'ISD_DISTRIBUTOR_STATE', 'ISD_DISTRIBUTOR_PINCODE',
                    'ISD_DISTRIBUTOR_STATE_CODE', 'CREDIT_RECIPIENT_GSTIN',
                    'CREDIT_RECIPIENT_NAME', 'CREDIT_RECIPIENT_ADDRESS',
                    'CREDIT_RECIPIENT_STATE', 'CREDIT_RECIPIENT_PINCODE',
                    'CREDIT_RECIPIENT_STATE_CODE',
                    'ELIGIBLE_CGST', 'ELIGIBLE_SGST', 'ELIGIBLE_UTGST', 'ELIGIBLE_IGST',
                    'INELIGIBLE_CGST', 'INELIGIBLE_SGST', 'INELIGIBLE_UTGST', 'INELIGIBLE_IGST',
                    'AMOUNT', 'REG_OFFICE', 'CIN', 'E_MAIL', 'WEBSITE'
                ],
                dtype=str,
                na_values=['', 'NA', 'N/A', 'NULL'],
                keep_default_na=False
            )
            logging.info(f"Successfully loaded Excel file: {file_path}")

        elif file_ext == '.csv':
            # Existing CSV handling
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