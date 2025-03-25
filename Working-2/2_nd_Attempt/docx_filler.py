import os
import re
import logging
from docx import Document
import pandas as pd
from docx.shared import Pt
from typing import Dict, List, Optional
from num2words import num2words  # For amount-to-words conversion

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def fill_docx_template(template_path, output_path, replacements):
    try:
        doc = Document(template_path)

        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for key, value in replacements.items():
                if f"{{{{{key}}}}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))

        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in replacements.items():
                        if f"{{{{{key}}}}}" in cell.text:
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))

        doc.save(output_path)
        return True
    except Exception as e:
        logging.error(f"Error filling template: {str(e)}")
        return False

        os.makedirs(output_folder, exist_ok=True)
        generated_files = []

        for idx, row in data.iterrows():
            try:
                doc = Document(template_path)
                row_data = process_row(row, amount_columns, convert_amount_to_words)

                if not replace_placeholders_in_document(doc, row_data):
                    continue

                output_path = save_document(doc, output_folder, idx)
                if output_path:
                    generated_files.append(output_path)

            except Exception as e:
                logging.error(f"Error processing row {idx + 1}: {str(e)}")
                continue

        return generated_files if generated_files else None

    except Exception as e:
        logging.error(f"Fatal error: {str(e)}")
        return None


def validate_inputs(template_path: str, data: pd.DataFrame, output_folder: str) -> bool:
    """Validate all input parameters"""
    if not os.path.exists(template_path):
        logging.error(f"Template file not found: {template_path}")
        return False

    if data.empty:
        logging.error("No data provided in DataFrame")
        return False

    return True

def convert_amount_to_words(amount: float) -> str:
    """
    Converts numeric amount to words representation.
    Example: 1234.56 → "One Thousand Two Hundred Thirty-Four Point Five Six"

    Args:
        amount: Numeric value to convert

    Returns:
        String representation in words
    """
    try:
        if pd.isna(amount):
            return ""

        # Split into dollars and cents
        dollars = int(amount)
        cents = round((amount - dollars) * 100)

        dollar_words = num2words(dollars, lang='en').title()

        if cents > 0:
            cent_words = num2words(cents, lang='en').title()
            return f"{dollar_words} And {cent_words} Cents"
        return f"{dollar_words} Only"

    except Exception as e:
        logging.warning(f"Amount-to-words conversion failed: {str(e)}")
        return ""


def process_row(
        row: pd.Series,
        amount_columns: Optional[List[str]],
        convert_to_words: bool = False
) -> Dict:
    """Process row data with enhanced amount handling"""
    row_data = row.to_dict()

    # Handle Amount field
    if 'Amount' in row_data:
        if pd.isna(row_data['Amount']) or is_formula(row_data['Amount']):
            if amount_columns:
                row_data['Amount'] = sum_numeric_columns(row, amount_columns)
            else:
                row_data['Amount'] = sum_all_numeric_columns(row)

        # Add amount in words if requested
        if convert_to_words:
            row_data['Amount_In_Words'] = convert_amount_to_words(row_data['Amount'])

    # Format all values
    return {k: format_value(v) for k, v in row_data.items()}


def is_formula(value) -> bool:
    """Check if a value might be an Excel formula"""
    return isinstance(value, str) and value.startswith('=')


def sum_numeric_columns(row: pd.Series, columns: List[str]) -> float:
    """Sum specified numeric columns"""
    try:
        return sum(float(row[col]) for col in columns if pd.notna(row.get(col)))
    except (ValueError, TypeError):
        logging.warning(f"Couldn't sum columns {columns}")
        return 0.0


def sum_all_numeric_columns(row: pd.Series) -> float:
    """Sum all numeric columns in the row"""
    try:
        return sum(v for v in row if isinstance(v, (int, float)))
    except TypeError:
        return 0.0


def format_value(value) -> str:
    """Format values for document insertion"""
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)):
        return "{:,.2f}".format(value)

    return str(value)


def replace_placeholders_in_document(doc: Document, row_data: Dict) -> bool:
    """Replace placeholders throughout document components"""
    try:
        # Normalize row_data keys to lowercase for case-insensitive matching
        normalized_data = {str(k).lower(): str(v) for k, v in row_data.items()}

        # Process all document parts
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, normalized_data)

        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph, normalized_data)

        # Process headers and footers
        for section in doc.sections:
            for header in [section.header, section.first_page_header]:
                if header:
                    for paragraph in header.paragraphs:
                        replace_in_paragraph(paragraph, normalized_data)

            for footer in [section.footer, section.first_page_footer]:
                if footer:
                    for paragraph in footer.paragraphs:
                        replace_in_paragraph(paragraph, normalized_data)

        return True
    except Exception as e:
        logging.error(f"Error replacing placeholders: {str(e)}", exc_info=True)
        return False


def replace_in_paragraph(paragraph, row_data: Dict):
    """Replace placeholders in a paragraph with actual values"""
    if not paragraph.text.strip():
        return

    # Combine runs to handle split placeholders
    full_text = ''.join(run.text for run in paragraph.runs)
    original_text = full_text

    # Find all placeholders in the text
    placeholders_in_text = re.findall(r'\{\{\s*(.*?)\s*\}\}', full_text, re.IGNORECASE)

    # Replace each found placeholder
    for ph in set(placeholders_in_text):  # Use set to avoid duplicates
        ph_lower = ph.lower()
        if ph_lower in row_data:
            full_text = full_text.replace(f'{{{{{ph}}}}}', row_data[ph_lower])

    # Only modify if changes were made
    if full_text != original_text:
        paragraph.clear()
        if full_text.strip():  # Only add run if there's content
            paragraph.add_run(full_text)

def save_document(doc: Document, output_folder: str, idx: int) -> Optional[str]:
    """Save filled document with sequential naming"""
    try:
        filename = f"document_{idx + 1}.docx"
        output_path = os.path.join(output_folder, filename)
        doc.save(output_path)
        logging.info(f"Generated: {output_path}")
        return output_path
    except Exception as e:
        logging.error(f"Error saving document {idx + 1}: {str(e)}")
        return None


def extract_placeholders(template_path: str) -> set:
    """
    Extracts all placeholders from a DOCX template, including headers/footers.

    Args:
        template_path: Path to the DOCX template file

    Returns:
        Set of cleaned placeholder names (e.g., {'invoice_number', 'amount'})

    Raises:
        FileNotFoundError: If template doesn't exist
        ValueError: If template is invalid
    """
    try:
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")

        doc = Document(template_path)
        placeholders = set()
        placeholder_pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")  # Handles whitespace

        def extract_from_text(text: str):
            """Inner function to extract placeholders from text"""
            return {match.strip().lower() for match in placeholder_pattern.findall(text)}

        # Process main document paragraphs (including formatted runs)
        for paragraph in doc.paragraphs:
            placeholders.update(extract_from_text(paragraph.text))
            for run in paragraph.runs:
                placeholders.update(extract_from_text(run.text))

        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        placeholders.update(extract_from_text(paragraph.text))
                        for run in paragraph.runs:
                            placeholders.update(extract_from_text(run.text))

        # Process headers and footers
        for section in doc.sections:
            for header in [section.header, section.first_page_header]:
                if header:
                    for paragraph in header.paragraphs:
                        placeholders.update(extract_from_text(paragraph.text))

            for footer in [section.footer, section.first_page_footer]:
                if footer:
                    for paragraph in footer.paragraphs:
                        placeholders.update(extract_from_text(paragraph.text))

        # Clean results - remove empty matches and normalize
        return {ph for ph in placeholders if ph}  # Remove empty strings

    except Exception as e:
        raise ValueError(f"Failed to extract placeholders: {str(e)}")