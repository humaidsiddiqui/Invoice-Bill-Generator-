from openpyxl import Workbook
import os

#create a new worksheet
wb = Workbook()

#activate a new worksheet
wb.active.title="Salary"

#another method to create a sheet
new_Sheet=wb.create_sheet("salary")

#rename an existing sheet
new_Sheet.title="renamedsheet"


#copy a sheet 
copied_sheet= wb.copy_worksheet(new_Sheet)
copied_sheet.title="CopiedSheet"


wb.save('otP/03.access_manage_worksheet.xlsx')
print("workbook saved successfully/")

