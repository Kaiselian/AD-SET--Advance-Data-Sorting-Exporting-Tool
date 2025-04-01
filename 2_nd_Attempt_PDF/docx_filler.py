import os
import re
import logging
from docx import Document
from docx.shared import Pt
from typing import Dict, List, Optional
from datetime import datetime
from data_mapper import prepare_row_data, map_data_to_docx


logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def fill_docx_template(template_path: str, output_path: str, replacements: Dict) -> bool:
    """
    Fill a DOCX template with provided replacements while preserving formatting
    Args:
        template_path: Path to the template DOCX file
        output_path: Path to save the filled document
        replacements: Dictionary of placeholder-value pairs
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        doc = Document(template_path)

        if not replace_all_placeholders(doc, replacements):
            return False

        doc.save(output_path)
        return True

    except Exception as e:
        logging.error(f"Error filling template: {str(e)}")
        return False


def replace_all_placeholders(doc: Document, row_data: Dict[str, str]) -> bool:
    """
    Replace placeholders throughout all document components
    Args:
        doc: The Document object to process
        row_data: Dictionary of placeholder replacements
    Returns:
        bool: True if successful, False if errors occurred
    """
    try:
        # Process main document paragraphs
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, row_data)

        # Process tables
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
    """
    Replace placeholders in a paragraph while preserving formatting
    Handles bold formatting (**text**) in the template
    Args:
        paragraph: The paragraph to process
        row_data: Dictionary of placeholder replacements
    """
    # Combine runs to handle split placeholders
    full_text = ''.join(run.text for run in paragraph.runs)

    # Skip if no placeholders or bold markers
    if not (any(f'{{{{{ph}}}}}' in full_text for ph in row_data) or '**' in full_text):
        return

    # Perform all placeholder replacements
    modified_text = full_text
    for ph, value in row_data.items():
        modified_text = modified_text.replace(f'{{{{{ph}}}}}', str(value))

    # Only update if changes were made
    if modified_text != full_text:
        paragraph.clear()

        # Split text by bold markers and process each segment
        parts = modified_text.split('**')
        for i, part in enumerate(parts):
            run = paragraph.add_run(part)
            run.font.size = Pt(10)  # Maintain font size

            # Apply bold to text between ** markers
            if i % 2 == 1:
                run.bold = True

            # Preserve original font if available
            if paragraph.runs and paragraph.runs[0].font.name:
                run.font.name = paragraph.runs[0].font.name


def generate_output_filename(row_data: Dict, idx: int, is_eligible: bool) -> str:
    """
    Generate a standardized output filename
    Args:
        row_data: Dictionary containing row data
        idx: Row index
        is_eligible: Whether this is an eligible invoice
    Returns:
        str: Generated filename
    """
    invoice_num = str(row_data.get('INVOICE_NUMBER', idx + 1)).strip()
    prefix = "Eligible" if is_eligible else "Ineligible"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_ISD_{invoice_num}_{timestamp}.docx"


def validate_template(template_path: str) -> bool:
    """
    Validate that the template exists and is accessible
    Args:
        template_path: Path to the template file
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        if not os.path.exists(template_path):
            logging.error(f"Template file not found: {template_path}")
            return False
        # Try opening the document to verify it's valid
        Document(template_path)
        return True
    except Exception as e:
        logging.error(f"Invalid template file: {str(e)}")
        return False