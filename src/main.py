from openpyxl import *

wb = load_workbook("resources\\ІТтаКБ. Сем I. Форма навчання  денна.xlsx")
ws = wb["Робота кафедри"]
wcell1 = ws.cell(20, 2)
wcell1.value = "PYTHON HERNYA"
wb.save("resources\\ІТтаКБ. Сем I. Форма навчання  денна.xlsx")
