import tabula

pdf_path = input("Enter the path or URL of the PDF file: ")

dfs = tabula.read_pdf(pdf_path, pages='all')

for i, df in enumerate(dfs):
    print(f"Table {i}:")
    print(df.head()) 
    print()
    dfs[i].to_csv(f"table_{i}.csv")

table_index = int(input("Enter the index of the table you want to select rows and columns from: ")) - 1

columns = input("Enter the column names (comma-separated) you want to select: ").split(",")

rows = input("Enter the row indices (comma-separated) you want to select: ").split(",")
rows = [int(row) for row in rows]

selected_df = dfs[table_index].loc[rows, columns]

selected_df.to_csv("selected_table.csv", index=False)

print(selected_df)

'''import tabula

pdf_path = input("Enter the path or URL of the PDF file: ")

dfs = tabula.read_pdf(pdf_path, pages='all')

for i, df in enumerate(dfs):
    print(f"Table {i + 1}:")
    print(df.head()) 
    print()

table_index = int(input("Enter the index of the table you want to select rows and columns from: ")) - 1

columns = input("Enter the column names (comma-separated) you want to select: ").split(",")

rows = input("Enter the row indices (comma-separated) you want to select: ").split(",")
rows = [int(row) for row in rows]

selected_df = dfs[table_index].loc[rows, columns]

selected_df.to_csv("selected_table.csv", index=False)

print(selected_df)'''