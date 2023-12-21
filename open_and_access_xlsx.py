# import the workbook 'wb' object from load_workbook
from loadWorkbook.load_workbook import wb


# wb object has attribute sheetnames that prints the name of the sheets in the excel file
print(wb.sheetnames)

# Access the individual cell:
# First access to the Sheet1 and it returns the sheet object
sheet = wb["Sheet1"]

# Now get access the cell column A row1, means a1
cell = sheet["a1"]

# Print the cell value
print(cell.value)

# Change the value for the cell:
cell.value = "Transaction_id"
# wb.save("saved_value.xlsx")

# Other attributes of cell object and print them in terminal:
print(cell.row)
print(cell.column)
print(cell.coordinate)

# Another approach to access the cell:
cell = sheet.cell(row=1, column=1)
print(cell.value)
print(cell.row)
print(cell.column)
print(cell.coordinate)

# Access various cells by using max_row and max_column and print them:
print(sheet.max_row)
print(sheet.max_column)

# OutPut: 4 (that means this sheet has 4 rows) and 
#         3 (That means this sheet has 3 columns)

# Now iterate over the the rows and colums to get all the values of all the cells and colums
for row in range(1, sheet.max_row + 1):
  for column in range(1, sheet.max_column + 1):
    cells = sheet.cell(row, column)

    print(cells.value)

# Access all the rows in a column using square brackets:
cells_in_column_A = sheet["a"]
print(cells_in_column_A)

# Now we cam access all the rows in column A to C by using square brackets:
cells_in_column_a_to_c = sheet["a:c"]
print(cells_in_column_a_to_c)

# Also use corridate here form a1 to c3 that returns all the area on column a to c row 1 to 3:
coordinate = sheet["a1:c3"]
print(coordinate)

# Also access cells in given rows:
cells_in_row_1 = sheet[1]
print(cells_in_row_1)

# Access cells in given row in range like row 1 to 3:
cells_in_row_1_to_3 = sheet[1:3]
print(cells_in_row_1_to_3)
