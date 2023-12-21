# Import wb object which has a xlsx file from loadWorkbook package
from loadWorkbook.load_workbook import wb

# Get the sheet's name of the given xlsx file
print(wb.sheetnames)

# Access the Sheet1 and it returns name object
rows_columns = wb["Sheet1"]

# The sheet has 3 coulmns and 4 rows and we append 1, 2, 3 in row 5
# We use append() method and takes a list of tuple e.g ([1, 2, 3]) as argument
rows_columns.append([1, 2, 3])

# Then save the changes in the new worbook
wb.save("updated_transactions.xlsx")
