import fitz  # PyMuPDF


def extract_placeholders_from_pdf(pdf_path):
    """
    Extracts all placeholders {{...}} from a readable text PDF.

    :param pdf_path: Path to the PDF file.
    :return: Set of detected placeholders.
    """
    try:
        doc = fitz.open(pdf_path)
        placeholders = set()

        for page_num, page in enumerate(doc):
            text = page.get_text("text")  # Extract text from the page
            detected = {word for word in text.split() if word.startswith("{{") and word.endswith("}}")}
            placeholders.update(detected)

        if placeholders:
            print("✅ Placeholders found in the PDF:")
            for ph in placeholders:
                print(f"   ➜ {ph}")
        else:
            print("⚠️ No placeholders detected! Ensure they are in a selectable text format.")

        return placeholders

    except Exception as e:
        print(f"❌ Error reading PDF: {e}")
        return None


# Example Usage
pdf_template = "LensKart Invoice-Template.pdf"
placeholders = extract_placeholders_from_pdf(pdf_template)
