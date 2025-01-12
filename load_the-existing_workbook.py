from openpyxl import load_workbook
import  os


#load an existing workbook
lw=load_workbook("otP\crreatingandsaving_workbook.xlsx")

print(f"loaded sheet title: {lw.active.title}")
