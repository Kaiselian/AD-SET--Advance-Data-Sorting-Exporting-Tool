import os
import logging
from docx2pdf import convert

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

def generate_pdfs(docx_files, output_folder=None):
    """
    Converts DOCX files to PDFs and saves them in the specified output folder.

    Args:
        docx_files (list): List of paths to the DOCX files.
        output_folder (str): Path to the output folder (default: same as DOCX files).

    Returns:
        list: List of paths to the generated PDF files, or an empty list if an error occurs.
    """
    pdf_files = []

    try:
        # Validate inputs
        if not docx_files:
            logger.error("No DOCX files provided.")
            return pdf_files

        # Create output folder if specified
        if output_folder and not os.path.exists(output_folder):
            logger.info(f"Creating output folder: {output_folder}")
            os.makedirs(output_folder)

        # Convert each DOCX file to PDF
        for docx_file in docx_files:
            try:
                # Validate DOCX file
                if not os.path.exists(docx_file):
                    logger.error(f"DOCX file not found: {docx_file}")
                    continue
                if not docx_file.endswith(".docx"):
                    logger.error(f"Invalid file format: {docx_file}. Expected a .docx file.")
                    continue

                # Determine output path
                if output_folder:
                    pdf_output = os.path.join(output_folder, os.path.basename(docx_file).replace(".docx", ".pdf"))
                else:
                    pdf_output = docx_file.replace(".docx", ".pdf")

                # Convert DOCX to PDF
                logger.info(f"Converting {docx_file} to PDF...")
                convert(docx_file, pdf_output)
                pdf_files.append(pdf_output)
                logger.info(f"Saved PDF: {pdf_output}")

            except Exception as e:
                logger.error(f"Error converting {docx_file}: {e}")

        logger.info(f"Successfully converted {len(pdf_files)} files to PDF.")
        return pdf_files

    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        return pdf_files