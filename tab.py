import tabula
import pandas as pd
from openpyxl import load_workbook

pdf_path = "sample3.pdf"
excel_path = "out3.xlsx"

dfs = tabula.read_pdf(pdf_path, pages='all')

with pd.ExcelWriter(excel_path) as writer:
    for i, df in enumerate(dfs):
        df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

excel_path = "out3.xlsx"
workbook = load_workbook(excel_path)
