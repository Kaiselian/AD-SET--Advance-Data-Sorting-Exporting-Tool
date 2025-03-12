import os
from docx import Document
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

def replace_placeholder_in_runs(paragraph, placeholder, replacement):
    """
    Replaces placeholders in a paragraph's runs while preserving formatting.

    Args:
        paragraph: A docx paragraph object.
        placeholder (str): The placeholder text to replace.
        replacement (str): The text to replace the placeholder with.
    """
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)

def fill_docx_template(template_path, data, output_folder, file_prefix="filled"):
    """
    Fills a DOCX template for each row in the dataset and saves the filled documents.

    Args:
        template_path (str): Path to the DOCX template.
        data (pd.DataFrame): DataFrame containing the data.
        output_folder (str): Path to the output folder.
        file_prefix (str): Prefix for the output file names (default: "filled").

    Returns:
        list: List of paths to the filled DOCX files, or an empty list if an error occurs.
    """
    filled_files = []

    try:
        # Validate inputs
        if not os.path.exists(template_path):
            logger.error(f"Template file not found: {template_path}")
            return filled_files
        if not os.path.exists(output_folder):
            logger.info(f"Creating output folder: {output_folder}")
            os.makedirs(output_folder)
        if data.empty:
            logger.error("The DataFrame is empty.")
            return filled_files

        # Process each row in the DataFrame
        for idx, row in data.iterrows():
            try:
                # Load the template
                doc = Document(template_path)

                # Replace placeholders in each paragraph
                for para in doc.paragraphs:
                    for col in data.columns:
                        placeholder = f"{{{{{col.strip()}}}}}"
                        if placeholder in para.text:
                            replace_placeholder_in_runs(para, placeholder, str(row[col]))

                # Save the filled document
                output_path = os.path.join(output_folder, f"{file_prefix}_{idx + 1}.docx")
                doc.save(output_path)
                filled_files.append(output_path)
                logger.info(f"Saved filled document: {output_path}")

            except Exception as e:
                logger.error(f"Error processing row {idx + 1}: {e}")

        logger.info(f"Successfully filled {len(filled_files)} documents.")
        return filled_files

    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        return filled_files