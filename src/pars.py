from openpyxl.styles import *
from openpyxl import *
from openpyxl.worksheet.worksheet import Worksheet


class pars:
    def __init__(self):
        wb = load_workbook("resources\\Бакалавр (Денна)\\РНП 122 КН КН 1 курс скороч. 2р. 2019-2020 (2019-11-19).xlsx")
        wd = load_workbook("resources\\ІТтаКБ. Сем I. Форма навчання  денна.xlsx")
        sheet: Worksheet = wb.active
        sheett: Worksheet = wd.active
        print("DENNA")
        print(wb.get_sheet_names())
        c = "sd"
        a = []
        i = 0
        j = 0

        for i in range(1, sheet.max_row):
            for j in range(1, sheett.max_row):
                if sheett.cell(row=j, column=2).value == sheet.cell(row=i, column=2).value and sheett.cell(row=j, column=2).value != None:
                    if sheet.cell(row=i, column=32).value == sheett.cell(row=j, column=13).value:
                        print(i, sheet.cell(row=i, column=2).value)
                        print(j, sheett.cell(row=j, column=2).value)
                    elif sheet.cell(row=i, column=32).value == None and sheett.cell(row=j, column=13).value == "--":
                        print(i, sheet.cell(row=i, column=2).value, sheet.cell(row=i, column=24).value, sheet.cell(row=i, column=26).value, sheet.cell(row=i, column=28).value)
                        print(j, sheett.cell(row=j, column=2).value, sheett.cell(row=j, column=10).value, sheett.cell(row=j, column=11).value, sheett.cell(row=j, column=12).value)
                        if sheet.cell(row=i, column=24).value != sheett.cell(row=j, column=10).value:
                            sheet.cell(row=i, column=24).fill = PatternFill(fill_type='solid', start_color='ff8327', end_color='ff8327')  # Данный код позволяет делать оформление цветом ячейки

                        if sheet.cell(row=i, column=26).value != sheett.cell(row=j, column=11).value:
                            sheet.cell(row=i, column=26).fill = PatternFill(fill_type='solid', start_color='ff8327', end_color='ff8327')  # Данный код позволяет делать оформление цветом ячейки

                        if sheet.cell(row=i, column=28).value != sheett.cell(row=j, column=12).value:
                            sheet.cell(row=i, column=28).fill = PatternFill(fill_type='solid', start_color='ff8327', end_color='ff8327')  # Данный код позволяет делать оформление цветом ячейки

        wb.save("resources\\Бакалавр (Денна)\\РНП 122 КН КН 1 курс скороч. 2р. 2019-2020 (2019-11-19).xlsx")
        print("SRABOTALO")
        wb.close()
        wd.close()
