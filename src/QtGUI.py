from PyQt5 import QtWidgets

from src.Doc1 import doc1
from src.convert import ConverterXLS
from src.design import Ui_MainWindow
from src.pars import pars


class AppGUI(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.load_file_rnp)
        self.pushButton_2.clicked.connect(self.load_file)
        self.pushButton_5.clicked.connect(self.compare_files)
        self.pushButton_8.clicked.connect(self.convert_file)
        self.pushButton_6.clicked.connect(self.create_report)
        self.only_file_name = []
        self.only_file_name_rnp = []
        self.default_file_names = []
        self.default_file_names_rnp = []
        self.radioButton.click()

    def load_file(self):
        files_name = QtWidgets.QFileDialog.getOpenFileNames(self, "Open File", "resources", "XLS files (*.xls *.xlsx)")
        files_name = files_name[0]

        if len(files_name):
            try:
                self.comboBox.clear()
                self.comboBox_2.clear()
                self.comboBox_5.clear()
                self.comboBox_4.clear()
                self.comboBox.setEditable(False)
                self.comboBox_2.setEditable(False)
                self.comboBox_5.setEditable(False)
                self.comboBox_4.setEditable(False)
                self.only_file_name = []
                temporary_array = []
                for item in range(0, len(files_name)):
                    temporary_array.append(files_name[item].replace('/', '\\', 10))

                files_name = temporary_array
                self.default_file_names = files_name

                for item in range(0, len(files_name)):
                    temporary_array2 = files_name[item].split('\\')
                    self.only_file_name.append(temporary_array2[-1])
                self.comboBox.addItems(self.only_file_name)
                self.comboBox_2.addItems(self.only_file_name)
                self.comboBox_5.addItems(self.only_file_name)
                self.comboBox_4.addItems(self.only_file_name)

            except FileNotFoundError:
                QtWidgets.QMessageBox("Open Source File", "Failed to read file\n'%s'")
            return

    def load_file_rnp(self):
        files_name = QtWidgets.QFileDialog.getOpenFileNames(self, "Open File", "resources", "XLS files (*.xls *.xlsx)")
        files_name = files_name[0]

        if len(files_name):
            try:
                self.comboBox_3.clear()
                self.comboBox_3.setEditable(False)
                self.only_file_name_rnp = []
                temporary_array = []
                for item in range(0, len(files_name)):
                    temporary_array.append(files_name[item].replace('/', '\\', 10))

                files_name = temporary_array
                self.default_file_names_rnp = files_name

                for item in range(0, len(files_name)):
                    temporary_array2 = files_name[item].split('\\')
                    self.only_file_name_rnp.append(temporary_array2[-1])
                self.comboBox_3.addItems(self.only_file_name_rnp)

            except FileNotFoundError:
                error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Open Source File",
                                              "Failed to read file\n", QtWidgets.QMessageBox.Ok)
                error.exec_()
            return

    def convert_file(self):
        try:
            ConverterXLS(self.default_file_names_rnp[self.comboBox_3.currentIndex()])
            self.default_file_names_rnp[self.comboBox_3.currentIndex()] += 'x'

            temp_array = list(self.comboBox_3.itemText(i) for i in range(self.comboBox_3.count()))
            print(self.default_file_names_rnp)

            temp_array[temp_array.index(self.comboBox_3.currentText())] += 'x'
            self.comboBox_3.clear()
            self.comboBox_3.addItems(temp_array)

        except IndexError:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Convert File", "Failed to convert file\n", QtWidgets.QMessageBox.Ok)
            error.exec_()
        return

    def compare_files(self):
        try:
            if self.radioButton.isChecked():
                pars_object = pars(self.default_file_names_rnp[self.comboBox_3.currentIndex()],
                                   self.default_file_names[self.comboBox.currentIndex()],
                                   self.default_file_names[self.comboBox_2.currentIndex()])
            else:
                pars_object = pars(self.default_file_names_rnp[self.comboBox_3.currentIndex()],
                                   self.default_file_names[self.comboBox_4.currentIndex()],
                                   self.default_file_names[self.comboBox_5.currentIndex()])
            info = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Compare file",
                                         "Compare succeeded!\n",
                                         QtWidgets.QMessageBox.Ok)
            info.exec_()
        except:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Convert File", "Failed to convert file\n",
                                          QtWidgets.QMessageBox.Ok)
            error.exec_()
        return

    def create_report(self):
        #try:
            save_name = QtWidgets.QFileDialog.getSaveFileName(self, "Save File", "resources",
                                                                "XLS files (*.xls *.xlsx)")
            report_Objcet = doc1(save_name, self.default_file_names[self.comboBox.currentIndex()],
                                 self.default_file_names[self.comboBox_2.currentIndex()],
                                 self.default_file_names[self.comboBox_4.currentIndex()],
                                 self.default_file_names[self.comboBox_5.currentIndex()])

            info = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Report file",
                                         "Report created!\n",
                                         QtWidgets.QMessageBox.Ok)
            info.exec_()
        #except:
         #   error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Report File", "Failed to create report file\n",
         #                                 QtWidgets.QMessageBox.Ok)
          #  error.exec_()
        #return
