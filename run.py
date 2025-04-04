from extractor import extract_tables_from_pdf

# Replace with your actual file
pdf_file = "test3.pdf"
output_excel = "output.xlsx"

extract_tables_from_pdf(pdf_file, output_excel)
print(f"✅ Tables extracted and saved to {output_excel}")
