def read_pdf(pdf_path, excel_path):
    dfs = tabula.read_pdf(pdf_path, pages='all')
    with pd.ExcelWriter(excel_path) as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

    workbook = load_workbook(excel_path)
    sheet_names = workbook.sheetnames

    total_sheets = len(sheet_names)

    # NCR, Material, Design (in order)
    cells_table1 = [[2,4], [4,1], [6,3]]
    total_cells_table1 = len(cells_table1)

    # (DefectType, CharNo)4, (SerialNos)5, (Description)6
    cells_table2 = [[1,3], [1,4], [2,1], [2,1]]
    total_cells_table2 = len(cells_table2)

    table1 = []

    empty_rows = total_sheets  # Number of rows
    empty_cols = 4  # Number of columns

    # Create an empty 2D array
    table2 = [[None for _ in range(empty_cols)] for _ in range(empty_rows)]

    i = 0
    j = 0 

    count = 0
    currentRow = 0
    select = 0

    for i in range(total_sheets):
        
        sheet = workbook[sheet_names[i]]

        if i == 0:
            for j in range(total_cells_table1):
                row_number = cells_table1[j][0]
                column_number = cells_table1[j][1]
                cell_value = sheet.cell(row=row_number, column=column_number).value
                data = cell_value
                if data != None:
                    table1.append(data)
                else:
                    table1.append("None")
            # print(table1)

        # print("------" + str(sheet) + "------")
        
        if i in range(2, total_sheets):

            if count == 0:
                count = count + 1
            
            elif count == 1:

                # print("====DefectType====")

                row_number = cells_table2[select][0]
                column_number = cells_table2[select][1]
                cell_value = sheet.cell(row=row_number, column=column_number).value
                data = cell_value
                # print(data[12:]) # + "(" + str(row_number) + ", " + str(column_number) + ")" + " Select = " + str(select))

                temp1 = data[12:]

                if temp1 != None:
                    table2[currentRow][0] = temp1
                else:
                    table2[currentRow][0] = "None" 
                

                select = select + 1

                # print("====CharNo====")

                row_number = cells_table2[select][0]
                column_number = cells_table2[select][1]
                cell_value = sheet.cell(row=row_number, column=column_number).value
                data = cell_value
                # print(data[11:]) # + "(" + str(row_number) + ", " + str(column_number) + ")" + " Select = " + str(select))

                temp2 = data[11:]

                if temp2 != None:
                    table2[currentRow][0] = temp2
                else:
                    table2[currentRow][0] = "None" 
                

                select = select + 1
                count = count + 1

            elif count == 2:

                # print("====SerialNos====")

                row_number = cells_table2[select][0]
                column_number = cells_table2[select][1]
                cell_value = sheet.cell(row=row_number, column=column_number).value
                data = cell_value
                # print(data) # + "(" + str(row_number) + ", " + str(column_number) + ")" + " Select = " + str(select))

                temp3 = data

                if temp3 != None:
                    table2[currentRow][0] = temp3
                else:
                    table2[currentRow][0] = "None"
                

                select = select + 1
                count = count + 1

            elif count == 3:

                # print("====Description====")
                
                row_number = cells_table2[select][0]
                column_number = cells_table2[select][1]
                cell_value = sheet.cell(row=row_number, column=column_number).value
                data = cell_value
                # print(data) # + "(" + str(row_number) + ", " + str(column_number) + ")" + " Select = " + str(select))

                temp4 = data
                
                if temp4 != None:
                    table2[currentRow][0] = temp4
                else:
                    table2[currentRow][0] = "None"

                select = 0
                count = count + 1

            elif count == 4:
                count = count + 1
                select = 0

            elif count == 5:
                # print("===========================================================")
                count = 0
                select = 0
                currentRow = currentRow + 1
    
    return [table1, table2, currentRow]
