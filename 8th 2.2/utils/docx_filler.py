from docx import Document
import pandas as pd
from docx2pdf import convert
import os
import re
import logging
from datetime import datetime
from num2words import num2words  # Import the num2words library

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def extract_placeholders(template_path):
    """
    Extracts all placeholders from a DOCX template.

    :param template_path: Path to the DOCX template file
    :return: Set of placeholders (e.g., {"invoice number", "invoice date", ...})
    """
    doc = Document(template_path)
    placeholders = set()

    # Regex to find placeholders like {{placeholder}}
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")

    # Search in paragraphs
    for para in doc.paragraphs:
        matches = placeholder_pattern.findall(para.text)
        placeholders.update(matches)

    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = placeholder_pattern.findall(cell.text)
                placeholders.update(matches)

    return placeholders

def convert_amount_to_words(row):
    """
    Converts the numeric value in the 'amount' column to words.

    :param row: A row from the DataFrame
    :return: Words representation of the amount (e.g., "One Thousand Two Hundred Thirty-Four")
    """
    try:
        amount = float(row["amount"])  # Get the numeric value from the "amount" column
        return num2words(amount, lang='en')  # Convert to words
    except (ValueError, TypeError):
        logging.warning(f"⚠️ WARNING: Invalid amount value: {row['amount']}")
        return ""

def fill_docx_template(template_path, df, output_folder):
    """
    Fills a DOCX template with data from a DataFrame and exports individual PDFs.

    :param template_path: Path to the DOCX template file
    :param df: Pandas DataFrame containing data
    :param output_folder: Folder to save the generated files
    :return: List of paths to the generated DOCX files
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)  # Ensure output directory exists

    # Extract placeholders from the template
    placeholders = extract_placeholders(template_path)
    logging.info(f"🔹 Found placeholders: {placeholders}")

    generated_files = []

    for index, row in df.iterrows():
        try:
            doc = Document(template_path)  # Load the template

            # Replace placeholders in paragraphs
            for para in doc.paragraphs:
                full_text = para.text  # Get the full text of the paragraph
                for placeholder in placeholders:
                    if placeholder in df.columns:
                        value = str(row[placeholder]) if pd.notna(row[placeholder]) else ""
                        full_text = full_text.replace(f"{{{{{placeholder}}}}}", value)
                    elif placeholder == "amount_in_words":
                        # Convert {{amount}} to words for {{amount_in_words}}
                        value = convert_amount_to_words(row)
                        full_text = full_text.replace(f"{{{{{placeholder}}}}}", value)
                # Update paragraph text
                para.text = full_text

            # Replace placeholders in tables
            for table in doc.tables:
                for table_row in table.rows:
                    for cell in table_row.cells:
                        for para in cell.paragraphs:
                            full_text = para.text  # Get the full text of the paragraph
                            for placeholder in placeholders:
                                if placeholder in df.columns:
                                    value = str(row[placeholder]) if pd.notna(row[placeholder]) else ""
                                    full_text = full_text.replace(f"{{{{{placeholder}}}}}", value)
                                elif placeholder == "amount_in_words":
                                    # Convert {{amount}} to words for {{amount_in_words}}
                                    value = convert_amount_to_words(row)
                                    full_text = full_text.replace(f"{{{{{placeholder}}}}}", value)
                            # Update paragraph text
                            para.text = full_text

            # Generate output filename dynamically
            template_name = os.path.splitext(os.path.basename(template_path))[0]  # Get template name without extension
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Add timestamp for uniqueness
            docx_filename = os.path.join(output_folder, f"{template_name}_Filled_{index + 1}_{timestamp}.docx")

            # Save filled DOCX
            doc.save(docx_filename)
            logging.info(f"✅ DOCX filled successfully! Saved as {docx_filename}")

            # Convert to PDF
            pdf_filename = os.path.join(output_folder, f"{template_name}_Filled_{index + 1}_{timestamp}.pdf")
            convert(docx_filename, pdf_filename)
            logging.info(f"✅ PDF generated successfully! Saved as {pdf_filename}")

            generated_files.append((docx_filename, pdf_filename))

        except Exception as e:
            logging.error(f"❌ Error processing row {index + 1}: {e}")

    return generated_files

# Example usage
if __name__ == "__main__":
    # Define paths
    template_path = "C:/Users/anich/Downloads/Tax Invoice - Example (2) (2).docx"
    data_path = "C:/Users/anich/Downloads/Fields.xlsx"
    output_folder = "C:/Users/anich/Downloads/Output"

    # Load data
    df = pd.read_excel(data_path)

    # Fill the template
    fill_docx_template(template_path, df, output_folder)