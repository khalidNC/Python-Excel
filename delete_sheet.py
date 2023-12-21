# Import load_workload module
import loadWorkbook.load_workbook as load_workbook


# Create a class object new_loader that takes the newly created file
new_loader = load_workbook.LoadWorkbook("home_sheet.xlsx")
# Call the loadWorkbook function to load the workbook
wb_new = new_loader.loadWorkbook()

# Print the sheet names of the newly created sheet
print(wb_new.sheetnames)

# Access the sheet named home 
home_sheet = wb_new["home"]

# Delete the home sheet
wb_new.remove(home_sheet)

# Save the changes in the new file
wb_new.save("deleted_home_sheet.xlsx")
