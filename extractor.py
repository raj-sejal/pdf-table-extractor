 import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook


def extract_text_layout(pdf_path):
    doc = fitz.open(pdf_path)
    pages_data = []

    for page in doc:
        blocks = page.get_text("blocks")  # x0, y0, x1, y1, "text"
        text_blocks = []

        for b in blocks:
            if b[4].strip():
                text_blocks.append({
                    'x0': b[0],
                    'y0': b[1],
                    'x1': b[2],
                    'y1': b[3],
                    'text': b[4].strip()
                })

        pages_data.append(text_blocks)

    return pages_data


def group_into_rows(blocks, y_tolerance=5):
    rows = []
    blocks = sorted(blocks, key=lambda b: b['y0'])

    current_row = []
    last_y = None

    for block in blocks:
        if last_y is None or abs(block['y0'] - last_y) <= y_tolerance:
            current_row.append(block)
            last_y = block['y0']
        else:
            rows.append(current_row)
            current_row = [block]
            last_y = block['y0']

    if current_row:
        rows.append(current_row)

    return rows


def build_table(rows):
    table = []

    for row in rows:
        sorted_row = sorted(row, key=lambda b: b['x0'])
        table.append([cell['text'] for cell in sorted_row])

    return table


def save_to_excel(tables, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            df = pd.DataFrame(table)
            df.to_excel(writer, index=False, header=False, sheet_name=f"Table_{i+1}")


def extract_tables_from_pdf(pdf_path, output_excel_path="output.xlsx"):
    pages_data = extract_text_layout(pdf_path)
    all_tables = []

    for page_blocks in pages_data:
        rows = group_into_rows(page_blocks)
        table = build_table(rows)
        all_tables.append(table)

    save_to_excel(all_tables, output_excel_path)
    return output_excel_path
