import os
import re
import pandas as pd
import logging
from typing import Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


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


def validate_row(row):
    """Validate individual row data for required fields and GSTIN format"""
    required_fields = ['INVOICE_NUMBER', 'INVOICE_DATE', 'ISD_DISTRIBUTOR_GSTIN']
    missing = [field for field in required_fields if pd.isna(row.get(field))]
    if missing:
        raise ValueError(f"Missing required fields: {missing}")

    # Validate GSTIN format
    gstin_fields = ['ISD_DISTRIBUTOR_GSTIN', 'CREDIT_RECIPIENT_GSTIN']
    for field in gstin_fields:
        if field in row and not re.match(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$', str(row[field])):
            logging.warning(f"Invalid GSTIN format in {field}: {row[field]}")


def read_excel_csv(file_path: str) -> Optional[pd.DataFrame]:
    try:
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            return None

        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext in ['.xlsx', '.xls']:
            # Read the Excel file with data_only=True to evaluate formulas
            df = pd.read_excel(
                file_path,
                engine='openpyxl',
                header=0,
                skiprows=[1],  # Skip the tax type labels row
                names=[
                    # Basic invoice info
                    'INVOICE_NUMBER', 'INVOICE_DATE',
                    # ISD Distributor info
                    'ISD_DISTRIBUTOR_GSTIN', 'ISD_DISTRIBUTOR_NAME',
                    'ISD_DISTRIBUTOR_ADDRESS', 'ISD_DISTRIBUTOR_STATE',
                    'ISD_DISTRIBUTOR_PINCODE', 'ISD_DISTRIBUTOR_STATE_CODE',
                    # Credit Recipient info
                    'CREDIT_RECIPIENT_GSTIN', 'CREDIT_RECIPIENT_NAME',
                    'CREDIT_RECIPIENT_ADDRESS', 'CREDIT_RECIPIENT_STATE',
                    'CREDIT_RECIPIENT_PINCODE', 'CREDIT_RECIPIENT_STATE_CODE',
                    # Eligible tax breakdown
                    'ELIGIBLE_IGST_AS_IGST', 'ELIGIBLE_CGST_AS_IGST',
                    'ELIGIBLE_SGST_AS_IGST', 'ELIGIBLE_IGST_SUM',
                    'ELIGIBLE_CGST_AS_CGST', 'ELIGIBLE_CGST_SUM',
                    'ELIGIBLE_SGST_UTGST_AS_SGST_UTGST', 'ELIGIBLE_SGST_UTGST_SUM',
                    'ELIGIBLE_AMOUNT',
                    # Ineligible tax breakdown
                    'INELIGIBLE_IGST_AS_IGST', 'INELIGIBLE_CGST_AS_IGST',
                    'INELIGIBLE_SGST_AS_IGST', 'INELIGIBLE_IGST_SUM',
                    'INELIGIBLE_CGST_AS_CGST', 'INELIGIBLE_CGST_SUM',
                    'INELIGIBLE_SGST_UTGST_AS_SGST_UTGST', 'INELIGIBLE_SGST_UTGST_SUM',
                    'INELIGIBLE_AMOUNT',
                    # Contact info
                    'REG_OFFICE', 'CIN', 'E_MAIL', 'WEBSITE'
                ],
                dtype=str,
                na_values=['', 'NA', 'N/A', 'NULL'],
                keep_default_na=False
            )

            # Convert numeric columns to float
            numeric_cols = [
                'ELIGIBLE_IGST_AS_IGST', 'ELIGIBLE_CGST_AS_IGST', 'ELIGIBLE_SGST_AS_IGST',
                'ELIGIBLE_IGST_SUM', 'ELIGIBLE_CGST_AS_CGST', 'ELIGIBLE_CGST_SUM',
                'ELIGIBLE_SGST_UTGST_AS_SGST_UTGST', 'ELIGIBLE_SGST_UTGST_SUM', 'ELIGIBLE_AMOUNT',
                'INELIGIBLE_IGST_AS_IGST', 'INELIGIBLE_CGST_AS_IGST', 'INELIGIBLE_SGST_AS_IGST',
                'INELIGIBLE_IGST_SUM', 'INELIGIBLE_CGST_AS_CGST', 'INELIGIBLE_CGST_SUM',
                'INELIGIBLE_SGST_UTGST_AS_SGST_UTGST', 'INELIGIBLE_SGST_UTGST_SUM', 'INELIGIBLE_AMOUNT'
            ]

            for col in numeric_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # Calculate any missing sums
            df['ELIGIBLE_IGST_SUM'] = df[
                ['ELIGIBLE_IGST_AS_IGST', 'ELIGIBLE_CGST_AS_IGST', 'ELIGIBLE_SGST_AS_IGST', 'ELIGIBLE_AMOUNT']].sum(axis=1)
            df['INELIGIBLE_IGST_SUM'] = df[
                ['INELIGIBLE_IGST_AS_IGST', 'INELIGIBLE_CGST_AS_IGST', 'INELIGIBLE_SGST_AS_IGST', 'INELIGIBLE_AMOUNT']].sum(axis=1)

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

        # Validate each row
        for idx, row in df.iterrows():
            try:
                validate_row(row)
            except ValueError as e:
                logging.error(f"Row {idx + 1} validation failed: {str(e)}")
                # Either remove invalid rows or raise exception
                # df.drop(index=idx, inplace=True)  # Option 1: Skip invalid rows
                raise ValueError(f"Row {idx + 1} invalid: {str(e)}")  # Option 2: Fail fast

        logging.info(f"Columns in data: {df.columns.tolist()}")
        logging.info(f"First row sample:\n{df.iloc[0].to_dict()}")

        return df

    except PermissionError:
        logging.error(f"Permission denied when reading: {file_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading {file_path}: {str(e)}")
        return None