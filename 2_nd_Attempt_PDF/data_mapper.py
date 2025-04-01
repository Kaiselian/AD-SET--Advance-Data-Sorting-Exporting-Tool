import os
import re
import logging
from docx import Document
import pandas as pd
from typing import List, Optional, Set, Dict
from datetime import datetime
from copy import deepcopy
from num2words import num2words
from docx.shared import Pt
from typing import Dict

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# Enhanced column mapping with both eligible and ineligible tax fields
COLUMN_MAPPING = {
    # Invoice fields
    'invoicenumber': 'INVOICE_NUMBER',
    'invoicedate': 'INVOICE_DATE',

    # ISD Distributor fields
    'isddistributorgstin': 'ISD_DISTRIBUTOR_GSTIN',
    'isddistributorname': 'ISD_DISTRIBUTOR_NAME',
    'isddistributoraddress': 'ISD_DISTRIBUTOR_ADDRESS',
    'isddistributorstate': 'ISD_DISTRIBUTOR_STATE',
    'isddistributorpincode': 'ISD_DISTRIBUTOR_PINCODE',
    'isddistributorstatecode': 'ISD_DISTRIBUTOR_STATE_CODE',

    # Credit Recipient fields
    'creditrecipientgstin': 'CREDIT_RECIPIENT_GSTIN',
    'creditrecipientname': 'CREDIT_RECIPIENT_NAME',
    'creditrecipientaddress': 'CREDIT_RECIPIENT_ADDRESS',
    'creditrecipientstate': 'CREDIT_RECIPIENT_STATE',
    'creditrecipientpincode': 'CREDIT_RECIPIENT_PINCODE',
    'creditrecipientstatecode': 'CREDIT_RECIPIENT_STATE_CODE',

    # Tax fields - Handle both eligible and ineligible
    'eligible cgst': 'ELIGIBLE_CGST',
    'eligible sgst': 'ELIGIBLE_SGST',
    'eligible utgst': 'ELIGIBLE_UTGST',
    'eligible igst': 'ELIGIBLE_IGST',
    'ineligible cgst': 'INELIGIBLE_CGST',
    'ineligible sgst': 'INELIGIBLE_SGST',
    'ineligible utgst': 'INELIGIBLE_UTGST',
    'ineligible igst': 'INELIGIBLE_IGST',
    'cgst': 'CGST',  # Fallback
    'sgst': 'SGST',  # Fallback
    'utgst': 'UTGST',  # Fallback
    'igst': 'IGST',  # Fallback

    # Amount fields
    'amount': 'AMOUNT',
    'total': 'AMOUNT',

    # Contact fields
    'regoffice': 'REG_OFFICE',
    'cin': 'CIN',
    'e-mail': 'E_MAIL',
    'website': 'WEBSITE',

    # Special fields
    'amount_in_words': 'AMOUNT_IN_WORDS'
}


def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Enhanced column name normalization"""
    df.columns = [
        col.strip().upper()
        .replace(' ', '_')
        .replace('-', '_')
        .replace('.', '')
        .replace('ELIGABLE', 'ELIGIBLE')  # Fix common typo
        for col in df.columns
    ]
    return df


def map_data_to_docx(template_path: str, data: pd.DataFrame, output_folder: str,
                    is_eligible: bool = True) -> Optional[List[str]]:
    """
    Main function to generate DOCX files with template selection
    Args:
        template_path: Path to the template file
        data: DataFrame containing the data
        output_folder: Output directory for generated files
        is_eligible: Boolean indicating whether to use eligible template
    """
    try:
        if not validate_inputs(template_path, data, output_folder):
            return None

        os.makedirs(output_folder, exist_ok=True)
        generated_files = []
        template_placeholders = scan_template_placeholders(template_path)

        logging.info(f"Processing {len(data)} rows with {'eligible' if is_eligible else 'ineligible'} template")

        for idx, row in data.iterrows():
            try:
                doc = Document(template_path)
                row_data = prepare_row_data(row, template_placeholders, is_eligible)

                if idx == 0:  # Debug info for first row
                    log_debug_info(row, template_placeholders, row_data)

                if not replace_all_placeholders(doc, row_data):
                    logging.error(f"Skipping row {idx} due to replacement errors")
                    continue

                output_path = generate_output_path(output_folder, row_data, idx, is_eligible)
                doc.save(output_path)
                generated_files.append(output_path)
                logging.info(f"Generated: {os.path.basename(output_path)}")

            except Exception as e:
                logging.error(f"Error processing row {idx}: {str(e)}", exc_info=True)
                continue

        return generated_files if generated_files else None

    except Exception as e:
        logging.error(f"Fatal error in document generation: {str(e)}", exc_info=True)
        return None


def validate_inputs(template_path: str, data: pd.DataFrame, output_folder: str) -> bool:
    """Validate all input parameters"""
    if not os.path.exists(template_path):
        logging.error(f"Template file not found: {template_path}")
        return False

    if data.empty:
        logging.error("No data provided in DataFrame")
        return False

    try:
        os.makedirs(output_folder, exist_ok=True)
        return True
    except Exception as e:
        logging.error(f"Output folder not writable: {str(e)}")
        return False


def prepare_row_data(row, template_placeholders, is_eligible):
    row_data = {}
    prefix = "ELIGIBLE_" if is_eligible else "INELIGIBLE_"

    # Process tax fields
    tax_types = ['CGST', 'SGST', 'UTGST', 'IGST']
    for tax in tax_types:
        col_name = prefix + tax
        value = safe_float_conversion(row[col_name]) if col_name in row else 0
        row_data[tax] = format_value(value, tax)

    # Calculate total amount
    total_amount = sum(float(row_data.get(tax, 0)) for tax in tax_types)
    row_data['Amount'] = format_value(total_amount, 'Amount')

    # Process amount in words
    if any('amount_in_words' in ph.lower() for ph in template_placeholders):
        try:
            words = num2words(total_amount, lang='en_IN').title()
            words = words.replace('And', 'and')
            row_data['amount_in_words'] = f"{words} Rupees Only"
        except Exception as e:
            logging.error(f"Amount to words conversion failed: {str(e)}")
            row_data['amount_in_words'] = ""

    # Map other fields
    field_mapping = {
        # Invoice fields
    'Invoice Number': 'INVOICE_NUMBER',
    'Invoice Date': 'INVOICE_DATE',

    # ISD Distributor fields
    'ISD Distributor GSTIN': 'ISD_DISTRIBUTOR_GSTIN',
    'ISD Distributor Name': 'ISD_DISTRIBUTOR_NAME',
    'ISD Distributor Address': 'ISD_DISTRIBUTOR_ADDRESS',
    'ISD Distributor State': 'ISD_DISTRIBUTOR_STATE',
    'ISD Distributor Pincode': 'ISD_DISTRIBUTOR_PINCODE',
    'ISD Distributor State Code': 'ISD_DISTRIBUTOR_STATE_CODE',

    # Credit Recipient fields
    'Credit Recipient GSTIN': 'CREDIT_RECIPIENT_GSTIN',
    'Credit Recipient Name': 'CREDIT_RECIPIENT_NAME',
    'Credit Recipient Address': 'CREDIT_RECIPIENT_ADDRESS',
    'Credit Recipient State': 'CREDIT_RECIPIENT_STATE',
    'Credit Recipient Pincode': 'CREDIT_RECIPIENT_PINCODE',
    'Credit Recipient State Code': 'CREDIT_RECIPIENT_STATE_CODE',

    # Tax fields - Handle both eligible and ineligible
    'Eligible cgst': 'ELIGIBLE_CGST',
    'Eligible sgst': 'ELIGIBLE_SGST',
    'Eligible utgst': 'ELIGIBLE_UTGST',
    'Eligible igst': 'ELIGIBLE_IGST',
    'Ineligible cgst': 'INELIGIBLE_CGST',
    'Ineligible sgst': 'INELIGIBLE_SGST',
    'Ineligible utgst': 'INELIGIBLE_UTGST',
    'Ineligible igst': 'INELIGIBLE_IGST',
    'cgst': 'CGST',  # Fallback
    'sgst': 'SGST',  # Fallback
    'utgst': 'UTGST',  # Fallback
    'igst': 'IGST',  # Fallback

    # Amount fields
    'Amount': 'AMOUNT',
    'Total': 'AMOUNT',

    # Contact fields
    'Reg. Office': 'REG_OFFICE',
    'CIN': 'CIN',
    'E-Mail': 'E_MAIL',
    'Website': 'WEBSITE',

    # Special fields
    'amount_in_words': 'AMOUNT_IN_WORDS'
    }

    for template_ph, data_col in field_mapping.items():
        if template_ph in template_placeholders:
            row_data[template_ph] = format_value(row.get(data_col, ''), template_ph)

    return row_data

def safe_float_conversion(value):
    """Safely convert values to float, handling various edge cases"""
    if pd.isna(value) or value in ['', None]:
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def replace_all_placeholders(doc: Document, row_data: Dict[str, str]) -> bool:
    """Replace placeholders throughout document with formatting preservation"""
    try:
        # Process all paragraphs in main document
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, row_data)

        # Process all tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph, row_data)

        # Process headers and footers
        for section in doc.sections:
            for header in [section.header, section.first_page_header]:
                if header:
                    for paragraph in header.paragraphs:
                        replace_in_paragraph(paragraph, row_data)

            for footer in [section.footer, section.first_page_footer]:
                if footer:
                    for paragraph in footer.paragraphs:
                        replace_in_paragraph(paragraph, row_data)

        return True

    except Exception as e:
        logging.error(f"Error replacing placeholders: {str(e)}", exc_info=True)
        return False


def replace_in_paragraph(paragraph, row_data: Dict[str, str]):
    """Enhanced placeholder replacement with bold formatting preservation"""
    # First combine all runs to handle split placeholders
    full_text = ''.join(run.text for run in paragraph.runs)

    # Skip if no placeholders or bold markers
    if not (any(f'{{{{{ph}}}}}' in full_text for ph in row_data) or '**' in full_text):
        return

    # Perform all placeholder replacements first
    modified_text = full_text
    for ph, value in row_data.items():
        modified_text = modified_text.replace(f'{{{{{ph}}}}}', str(value))

    # Only update if changes were made
    if modified_text != full_text:
        # Clear existing content
        paragraph.clear()

        # Split text by bold markers and process each segment
        parts = modified_text.split('**')
        for i, part in enumerate(parts):
            run = paragraph.add_run(part)
            run.font.size = Pt(10)

            # Apply bold to every odd segment (text between ** markers)
            if i % 2 == 1:  # This is text between ** markers
                run.bold = True

            # Preserve original font if available
            if paragraph.runs and paragraph.runs[0].font.name:
                run.font.name = paragraph.runs[0].font.name


def format_value(value, key=None) -> str:
    """Enhanced value formatting with special cases"""
    if pd.isna(value) or value in ['', None]:
        return ""

    # Handle numpy types
    if hasattr(value, 'item'):
        value = value.item()

    # Special formatting for amounts
    if key and 'amount' in key.lower() and isinstance(value, (int, float)):
        return "{:,.2f}".format(value)

    # Special handling for GSTIN (format with spaces)
    if key and 'gstin' in key.lower() and isinstance(value, str) and len(value) == 15:
        return f"{value[:2]} {value[2:5]} {value[5:7]} {value[7:12]} {value[12:15]}"

    return str(value).strip()


def scan_template_placeholders(template_path: str) -> Set[str]:
    """Extract all unique placeholders from a DOCX template"""
    doc = Document(template_path)
    placeholders = set()
    placeholder_pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")  # Handles whitespace

    def extract_from_text(text: str):
        return {match.strip() for match in placeholder_pattern.findall(text)}

    # Process all document components
    components = [
        doc.paragraphs,
        *[cell.paragraphs for table in doc.tables
          for row in table.rows
          for cell in row.cells],
        *[section.header.paragraphs for section in doc.sections],
        *[section.footer.paragraphs for section in doc.sections]
    ]

    for paragraphs in components:
        for paragraph in paragraphs:
            placeholders.update(extract_from_text(paragraph.text))
            for run in paragraph.runs:
                placeholders.update(extract_from_text(run.text))

    return {ph for ph in placeholders if ph}  # Remove empty strings


def generate_output_path(output_folder: str, row_data: dict, idx: int,
                         is_eligible: bool) -> str:
    """Generate output path with type prefix and invoice number"""
    invoice_num = str(row_data.get('INVOICE_NUMBER', idx + 1)).strip()
    prefix = "ELIGIBLE" if is_eligible else "INELIGIBLE"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_folder, f"{prefix}_ISD_{invoice_num}_{timestamp}.docx")


def log_debug_info(row, template_placeholders, row_data):
    """Enhanced debug logging with more details"""
    logging.info("\n=== DEBUG INFORMATION ===")
    logging.info(f"Template placeholders: {sorted(template_placeholders)}")
    logging.info(f"Data columns: {sorted(row.index.tolist())}")

    logging.info("\n=== PLACEHOLDER MAPPING ===")
    for ph in sorted(template_placeholders):
        norm_ph = ph.lower().replace(' ', '').replace('.', '').replace('-', '')
        data_key = COLUMN_MAPPING.get(norm_ph, "NO MATCH")
        logging.info(f"Template: {ph:25} → Data: {data_key}")

    logging.info("\n=== MATCHED DATA ===")
    for ph, value in sorted(row_data.items()):
        logging.info(f"{ph:25}: {value}")
    logging.info("=====================")