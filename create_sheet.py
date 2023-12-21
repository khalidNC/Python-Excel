# Import load_workbook module
import loadWorkbook.load_workbook as load_workbook

# Get sheet's name of the given xlsx file
print(load_workbook.wb.sheetnames)

# Create a new sheet named home on the second position
load_workbook.wb.create_sheet("home", 1)

# Save the sheet in a new xlss file
load_workbook.wb.save("home_sheet.xlsx")
