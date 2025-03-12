import os
from docx import Document
import pandas as pd

def extract_text_with_runs(paragraph):
    """Extracts text from a paragraph, including formatting splits."""
    text = ""
    for run in paragraph.runs:
        text += run.text
    return text

def map_data_to_docx(template_path, data, output_path="output_filled.docx"):
    print("🔹 DEBUG: Mapping Data to Template")
    print(f"Template Path: {template_path}")
    print("Columns in Data:", data.columns)

    try:
        doc = Document(template_path)
    except FileNotFoundError:
        print(f"❌ Error: Template file not found at {template_path}")
        return None
    except Exception as e:
        print(f"❌ Error: Could not open template file: {e}")
        return None

    placeholders_found = set()

    def replace_text_in_paragraphs(paragraphs, data):
        """Replace placeholders in paragraphs, preserving runs."""
        for paragraph in paragraphs:
            full_text = extract_text_with_runs(paragraph)
            for column in data.columns:
                placeholder = f"{{{{{column.strip()}}}}}"
                if placeholder in full_text:
                    print(f"✅ Found placeholder: {placeholder}")
                    try:
                        value = str(data[column].iloc[0])
                    except KeyError:
                        print(f"❌ Error: Column '{column}' not found in data.")
                        continue

                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)
                            placeholders_found.add(placeholder)
                else:
                    print(f"⚠️ WARNING: Placeholder {placeholder} NOT found in paragraph!")

    def replace_text_in_tables(tables, data):
        """Replace placeholders inside tables."""
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        full_text = extract_text_with_runs(paragraph)
                        for column in data.columns:
                            placeholder = f"{{{{{column.strip()}}}}}"
                            if placeholder in full_text:
                                print(f"✅ Found placeholder: {placeholder} in a table")
                                try:
                                    new_text = full_text.replace(placeholder, str(data[column].iloc[0]))
                                except KeyError:
                                    print(f"❌ Error: Column '{column}' not found in data.")
                                    continue
                                for run in paragraph.runs:
                                    run.clear()
                                paragraph.add_run(new_text)
                                placeholders_found.add(placeholder)

    replace_text_in_paragraphs(doc.paragraphs, data)
    replace_text_in_tables(doc.tables, data)

    for section in doc.sections:
        replace_text_in_paragraphs(section.header.paragraphs, data)
        replace_text_in_paragraphs(section.footer.paragraphs, data)

    try:
        doc.save(output_path)
        print(f"✅ DOCX filled successfully! Saved as {output_path}")
    except Exception as e:
        print(f"❌ Error: Could not save filled DOCX: {e}")
        return None

    if not placeholders_found:
        print("⚠️ WARNING: No placeholders were replaced. Check if placeholders match exactly in your DOCX.")

    return output_path