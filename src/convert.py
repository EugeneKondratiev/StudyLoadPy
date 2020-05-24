import win32com.client
fname = "C:\\Users\\GreamReaper\\PycharmProjects\\StudyLoadPy1\\src\\resources\\Бакалавр (Денна)\\РНП 122 КН КН 1 курс скороч. 2р. 2019-2020 (2019-11-19).xls"
excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()