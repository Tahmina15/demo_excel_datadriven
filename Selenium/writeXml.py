import openpyxl

path = "C:\\Users\\HP\\Documents\\Book1.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.get_sheet_by_name("Sheet2")

for r in range(1, 6):
    for c in range(1, 4):
        sheet.cell(row=r, column=c).value = "Welcome To Automation"

workbook.save(path)
