import os
import re
import logging
from docx import Document
import pandas as pd
from docx.shared import Pt
from typing import Dict, List, Optional, Set
from num2words import num2words

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def fill_docx_template(
        template_path: str,
        data: pd.DataFrame,
        output_folder: str,
        amount_columns: Optional[List[str]] = None,
        convert_amount_to_words: bool = True,
        font_size: int = 10
) -> Optional[List[str]]:
    """
    Enhanced DOCX template filler with complete features:
    - Handles all placeholder types (paragraphs, tables, headers/footers)
    - Preserves original formatting
    - Proper amount-to-words conversion
    - Font size consistency
    - Complete field handling without truncation

    Args:
        template_path: Path to DOCX template
        data: DataFrame containing replacement data
        output_folder: Output directory for generated files
        amount_columns: Columns to sum for Amount (optional)
        convert_amount_to_words: Whether to add amount in words
        font_size: Font size to enforce (default 10)

    Returns:
        List of generated file paths or None if error
    """
    try:
        if not validate_inputs(template_path, data, output_folder):
            return None

        os.makedirs(output_folder, exist_ok=True)
        generated_files = []
        template_placeholders = extract_placeholders(template_path)

        logging.info(f"Found placeholders in template: {template_placeholders}")
        logging.info(f"Columns in data: {data.columns.tolist()}")

        for idx, row in data.iterrows():
            try:
                doc = Document(template_path)
                row_data = prepare_row_data(row, template_placeholders, amount_columns, convert_amount_to_words)

                # Debug output for verification
                if idx == 0:  # Only print first row for debugging
                    logging.info(f"First row data mapping:\n{row_data}")

                if not replace_all_placeholders(doc, row_data, font_size):
                    logging.error(f"Skipping row {idx} due to replacement errors")
                    continue

                output_path = os.path.join(output_folder, f"Invoice_{idx + 1}.docx")
                doc.save(output_path)
                generated_files.append(output_path)
                logging.info(f"Generated: {output_path}")

            except Exception as e:
                logging.error(f"Error processing row {idx + 1}: {str(e)}", exc_info=True)
                continue

        return generated_files if generated_files else None

    except Exception as e:
        logging.error(f"Fatal error: {str(e)}", exc_info=True)
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
        # Test if we can create the output folder
        os.makedirs(output_folder, exist_ok=True)
        test_file = os.path.join(output_folder, 'test.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        return True
    except Exception as e:
        logging.error(f"Output folder not writable: {str(e)}")
        return False


def prepare_row_data(
        row: pd.Series,
        template_placeholders: Set[str],
        amount_columns: Optional[List[str]] = None,
        convert_to_words: bool = True
) -> Dict[str, str]:
    """
    Prepare complete row data with all required fields and proper formatting

    Args:
        row: Single row of data
        template_placeholders: Set of placeholders found in template
        amount_columns: Columns to sum for Amount (optional)
        convert_to_words: Whether to convert amount to words

    Returns:
        Dictionary of placeholder-value pairs
    """
    row_data = {}

    # First handle amount fields
    if 'Amount' in row:
        if pd.isna(row['Amount']) or is_formula(row['Amount']):
            if amount_columns:
                row['Amount'] = sum_numeric_columns(row, amount_columns)
            else:
                row['Amount'] = sum_all_numeric_columns(row)

    # Convert all values to proper strings
    for col in row.index:
        row_data[col] = format_value(row[col], col)

    # Handle special fields
    if 'amount_in_words' in template_placeholders and convert_to_words and 'Amount' in row:
        try:
            amount = float(row['Amount'])
            words = num2words(amount, lang='en_IN').title()
            words = words.replace(' And ', ' and ')  # Fix capitalization
            row_data['amount_in_words'] = f"{words} Rupees Only"
        except Exception as e:
            logging.warning(f"Amount-to-words conversion failed: {str(e)}")
            row_data['amount_in_words'] = ""

    # Ensure all template placeholders have values
    for ph in template_placeholders:
        if ph not in row_data:
            # Try to find matching column with flexible matching
            matched_col = None
            norm_ph = ph.lower().replace(' ', '_').replace('-', '_')

            for col in row.index:
                norm_col = col.lower().replace(' ', '_').replace('-', '_')
                if norm_col == norm_ph:
                    matched_col = col
                    break

            if matched_col:
                row_data[ph] = format_value(row[matched_col], ph)
            else:
                row_data[ph] = ""
                logging.warning(f"No data found for placeholder: {ph}")

    return row_data


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


def format_value(value, key=None) -> str:
    """
    Format values for document insertion with special handling

    Args:
        value: Value to format
        key: Optional column name for special formatting

    Returns:
        Formatted string value
    """
    if pd.isna(value):
        return ""

    # Handle numpy types
    if hasattr(value, 'item'):
        value = value.item()

    # Special formatting for different field types
    if isinstance(value, (int, float)):
        if key and 'amount' in key.lower():
            return "{:,.2f}".format(value)
        return str(value)

    # Special handling for GSTIN (format with spaces)
    if key and 'gstin' in key.lower() and isinstance(value, str) and len(value) == 15:
        return f"{value[:2]} {value[2:5]} {value[5:7]} {value[7:12]} {value[12:15]}"

    return str(value).strip()


def replace_all_placeholders(doc: Document, row_data: Dict[str, str], font_size: int) -> bool:
    """
    Replace placeholders throughout entire document while preserving formatting

    Args:
        doc: Document object
        row_data: Dictionary of placeholder-value pairs
        font_size: Font size to enforce

    Returns:
        True if successful, False if errors occurred
    """
    try:
        from docx.shared import Pt

        # Process all paragraphs in main document
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, row_data, font_size)

        # Process all tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph, row_data, font_size)

        # Process headers and footers
        for section in doc.sections:
            for header in [section.header, section.first_page_header]:
                if header:
                    for paragraph in header.paragraphs:
                        replace_in_paragraph(paragraph, row_data, font_size)

            for footer in [section.footer, section.first_page_footer]:
                if footer:
                    for paragraph in footer.paragraphs:
                        replace_in_paragraph(paragraph, row_data, font_size)

        return True

    except Exception as e:
        logging.error(f"Error replacing placeholders: {str(e)}", exc_info=True)
        return False


def replace_in_paragraph(paragraph, row_data: Dict[str, str], font_size: int):
    """
    Replace placeholders in a paragraph while preserving formatting

    Args:
        paragraph: Paragraph object
        row_data: Dictionary of placeholder-value pairs
        font_size: Font size to enforce
    """
    from docx.shared import Pt

    # First combine all runs to handle split placeholders
    full_text = ''.join(run.text for run in paragraph.runs)

    # Skip if no placeholders
    if not any(f'{{{{{ph}}}}}' in full_text for ph in row_data):
        return

    # Perform all replacements
    modified_text = full_text
    for ph, value in row_data.items():
        modified_text = modified_text.replace(f'{{{{{ph}}}}}', value)

    # Only update if changes were made
    if modified_text != full_text:
        # Clear existing content
        paragraph.clear()

        # Add new content with preserved formatting
        run = paragraph.add_run(modified_text)
        run.font.size = Pt(font_size)

        # Preserve other formatting from first run if available
        if paragraph.runs and paragraph.runs[0].font.name:
            run.font.name = paragraph.runs[0].font.name


def extract_placeholders(template_path: str) -> Set[str]:
    """
    Extract all unique placeholders from a DOCX template

    Args:
        template_path: Path to DOCX template file

    Returns:
        Set of cleaned placeholder names

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

    except Exception as e:
        raise ValueError(f"Failed to extract placeholders: {str(e)}")