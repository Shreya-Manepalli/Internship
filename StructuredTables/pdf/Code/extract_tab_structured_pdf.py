import PyPDF2
import tabula
import re
import pandas as pd

pdf_path = input("Enter the path or URL of the PDF file: ")

pdf_text = ""
with open(pdf_path, "rb") as file:
    pdf_reader = PyPDF2.PdfReader(file)
    num_pages = len(pdf_reader.pages)
    for page_number in range(num_pages):
        page = pdf_reader.pages[page_number]
        pdf_text += page.extract_text()

caption_pattern = (r'Table\s\d[:|;|.|\s][^\n]+')
captions = re.findall(caption_pattern, pdf_text)

dfs = tabula.read_pdf(pdf_path, pages='all', stream=True)

num_tables = len(dfs)

# Create an Excel writer
excel_writer = pd.ExcelWriter('tables.xlsx', engine='xlsxwriter')

for i, caption in enumerate(captions):
    print(f"{i + 1}: {caption.strip()}")

for i, df in enumerate(dfs):
    # Write each DataFrame to a separate sheet
    sheet_name = f"Table {i+1}"
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

excel_writer.save()

print("Tables have been copied to separate sheets in 'tables.xlsx' file.")

table_index = int(input("Enter the index of the table you want to select rows and columns from: ")) - 1

columns = input("Enter the column names (comma-separated) you want to select: ").split(",")

rows = input("Enter the row indices (comma-separated) you want to select: ").split(",")
rows = [int(row) for row in rows]

selected_df = dfs[table_index].loc[rows, columns]

selected_df.to_csv("selected_table.csv", index=False)

print(selected_df)
