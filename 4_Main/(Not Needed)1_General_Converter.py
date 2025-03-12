import pytesseract
from pdf2image import convert_from_path

# Path to your PDF
pdf_path = input("Enter Path: ")

# Convert PDF to images
images = convert_from_path(pdf_path)

# Extract text from each image
extracted_text = []
for img in images:
    text = pytesseract.image_to_string(img)
    extracted_text.append(text)

# Combine extracted text
full_text = "\n".join(extracted_text)
print(full_text)  # Print detected text
