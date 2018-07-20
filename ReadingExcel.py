#! python3

import openpyxl
import os

workbook = openpyxl.load_workbook('example.xlsx')   # loads the example.xlsx spreadsheet
print(type(workbook))                               # doublechecking that spreadsheet load works
sheet = workbook['Sheet1']                          # see all the sheets in the excel object

# You can pass a string of what cell you want with square bracket syntax
print(str(sheet['A1'].value))
print(str(sheet['B1'].value))
print(str(sheet['C1'].value))

# You can get info faster by calling the rows and columns. PS excel starts at 1 not 0
for i in range(1,8):
    print(i, sheet.cell(row=i, column=2).value)