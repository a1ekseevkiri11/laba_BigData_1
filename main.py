import openpyxl

wb = openpyxl.open("laba.xlsx")
table_1 = wb.worksheets[2]
table_2 = wb.worksheets[3]
print(table_1)
print(table_2)

