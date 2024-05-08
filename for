#Pythonのfor文でExcelの行削除

import openpyxl

wb = openpyxl.load_workbook("Sample.xlsx")
ws= wb.active

for i in range(1000, 1, -1):
    if ws.cell(row=i,column=1).value == 0:
        ws.delete_rows(i)

wb.save("Result.xlsx")
