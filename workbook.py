from openpyxl import load_workbook, Workbook

wb = load_workbook('Database.xlsx')
sheet = wb.active
a = sheet['A2'].value

print(a)