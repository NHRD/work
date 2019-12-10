import openpyxl as xls
from browser_control import browser_controller

items = browser_controller()

wb = xls.Workbook()
ws = wb.active
ws.title = "customar_list"

j = 1

for i in range(len(items)):
    cell = "B{}".format(j + 1)
    ws[cell] = items[i]
    j = j + 1
    
wb.save("customar_list.xlsx")