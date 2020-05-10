from openpyxl import *

wb = load_workbook("resources\\ІТтаКБ. Сем I. Форма навчання  денна.xlsx")
ws = wb["Робота кафедри"]
wcell1 = ws.cell(20, 2)
wcell1.value = "PYTHON HERNYA"
wcell2 = ws.cell(21, 2)
wcell2.value = "Пайтон Херня"
wb.save("resources\\ІТтаКБ. Сем I. Форма навчання  денна.xlsx")
