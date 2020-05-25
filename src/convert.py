import win32com.client

from PyQt5 import QtWidgets


class ConverterXLS:
    def __init__(self, file_name):
        if file_name != " " and file_name.find('.xlsx') == -1:
            excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(file_name)
            try:
                wb.SaveAs(file_name + 'x', FileFormat=51)  # FileFormat = 51 is for .xlsx extension
                info = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Convert excel file",
                                             "Convert xls to xlsx succeeded!\n",
                                             QtWidgets.QMessageBox.Ok)
                info.exec_()
            except:
                error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Save Source File",
                                              "You have declined convert this file\n",
                                              QtWidgets.QMessageBox.Ok)
                error.exec_()
            finally:
                wb.Close()  # FileFormat = 56 is for .xls extension
                excel.Application.Quit()

        else:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Open Source File",
                                          "This file is not xls\n",
                                          QtWidgets.QMessageBox.Ok)
            error.exec_()
