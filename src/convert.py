import win32com.client
from tkinter.messagebox import showerror, showinfo


class ConverterXLS:
    def __init__(self, file_name):
        if file_name != " " and file_name.find('.xlsx') == -1:
            excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(file_name)
            try:
                wb.SaveAs(file_name + 'x', FileFormat=51)  # FileFormat = 51 is for .xlsx extension
                showinfo('Convert excel file', 'Convert xls to xlsx succeeded!')
            except:
                showerror("Open Source File", "You have declined conver this file\n'%s'" % file_name)
            finally:
                wb.Close()  # FileFormat = 56 is for .xls extension
                excel.Application.Quit()

        else:
            showerror("Open Source File", "This file is not xls\n'%s'" % file_name)
