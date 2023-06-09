import tabula

pdf_path = "https://sedl.org/afterschool/toolkits/science/pdf/ast_sci_data_tables_sample.pdf"

dfs = tabula.read_pdf(pdf_path, pages='all')
for i in range(len(dfs)):
    dfs[i].to_csv(f"table_{i}.csv")