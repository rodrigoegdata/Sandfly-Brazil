import fitz
import re
import pandas as pd

def extract_lines_from_pdf(pdf_path, pattern):
    # Open the PDF file
    document = fitz.open(pdf_path)
    extracted_lines = []

    # Iterate through each page
    for page_num in range(document.page_count):
        # Extract text from the page
        page = document.load_page(page_num)
        text = page.get_text("text")

        # Find lines that match the pattern
        lines = text.split("\n")
        for line in lines:
            if re.search(pattern, line):
                extracted_lines.append(line)

    return extracted_lines

# Define the pattern to search for
pattern = r'BR\s?\(*?'

# Path to the PDF file
pdf_path = '\\Apostila-Vol.-1_2024.pdf'

# Extract lines from the PDF
lines_with_pattern = extract_lines_from_pdf(pdf_path, pattern)

# Path to save the Excel file
excel_path = 'especies_flebotomineos_brasil.xlsx'

def save_lines_to_excel(lines, excel_path):
    df = pd.DataFrame(lines, columns=['Extracted Lines'])
    df.to_excel(excel_path, index=False)

# Extract lines and save to Excel
lines = extract_lines_from_pdf(pdf_path, pattern)
save_lines_to_excel(lines, excel_path)
