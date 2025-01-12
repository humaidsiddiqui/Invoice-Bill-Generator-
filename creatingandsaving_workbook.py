from openpyxl import Workbook
import os

#crete a new workbook
wb = Workbook()


#select and activate the sheet
sheet = wb.active
sheet.title="my_sheet"



#ensurebthe directory exists
os.makedirs('otP', exist_ok=True)

#save the workbook
wb.save('otp/crreatingandsaving_workbook.xlsx')
print("WORKBOOK CREATED AND SAVED SUCCESFULLY")