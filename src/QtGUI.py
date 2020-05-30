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
        self.subjects = ['Дискретна математика',
                         'Додатково годин на',
                         'Години на ЕК',
                         'Дипломне проектування',
                         'Теорія інформації та кодування',
                         'Математична логіка',
                         'Переддипломна практика',
                         'Науково-дослідна практика',
                         'Структурне та логічне програмування',
                         'Технології комп\'ютерного проектування',
                         'Методи аналізу даних та їх організація',
                         'Методи та засоби комп\'ютерних iнформацiйних технологiй',
                         'Інформаційні технології управління виробництвом ',
                         'Технології розподілених систем та паралельних обчислень',
                         'Грід-системи та технології  хмарних обчислень',
                         'Комп\'ютерний та технічний дизайн',
                         'Комп\'ютерна графіка',
                         'Проектування в середовищі машинобудівних CAD',
                         'Геометричне моделювання та методи візуалізації',
                         'Алгоритмізація та програмування',
                         'Програмування',
                         'Алгоритми та методи обчислень та чисельні методи',
                         'Об\'єктно-орiєнтоване програмування',
                         'Проектно-технологічна практика',
                         'Алгоритми та методи обчислень',
                         'Методи та системи підтримки прийняття рішень',
                         'Розробка Web-систем',
                         'Моделювання систем',
                         'Чисельні методи',
                         'Методологія наукових досліджень',
                         'Веб технології та веб дизайн',
                         'Інформаційно-аналітична діяльність у галузі інформаційної безпеки',
                         'Інтелектуальний аналіз даних',
                         'Органiзацiя баз даних та знань',
                         'Методи та системи штучного iнтелекту',
                         'Зосередженi та розподiленi системи управлiння базами даних',
                         'Операційні системи',
                         'Управління IT проектами',
                         'Інформаційні технології',
                         'Експертні системи',
                         'Розробка програмного забезпечення',
                         'Інформатика і системологія',
                         'Основи автоматизованого проектування складних об`єктiв i систем',
                         'Інформатика та програмування',
                         'Крос-платформне програмування',
                         'Паралельне та багатопоточне програмування',
                         'Процедурне програмування',
                         'Сучасні системи управління базами даних ',
                         'Системний аналiз та проектування компютерних інформаційних систем',
                         'Моделі технології пректування та управління інформаційними системами',
                         'Чинники успішного працевлаштування',
                         'Технології захисту інформації',
                         'Теорія алгоритмів',
                         'Математичні методи дослідження операцій',
                         'Моделювання процесів та об\'єктів енергетичного машинобудування']
        print(len(self.subjects))
        self.lectors = ['Плотніков В. М.',
                        'Становська Т. П.',
                        'Мазурок І. Є.',
                        'Мазурок Т. Л.',
                        'Ломовцев П. Б.',
                        'Антонова А. Б.',
                        'Селіванова А. В.',
                        'Ольшевская',
                        'Котлик С. В.',
                        'Корнієнко Ю. К.',
                        'Тройнина',
                        'Сіромля С. Г.',
                        'Швець Н. В.',
                        'Попков Д. М.',
                        'Владімірова В. Б.',
                        'Бодюл О. С.']

        self.lectors.sort()
        self.subjects.sort()
        self.assisstents = ['Плотніков В. М.',
                        'Становська Т. П.',
                        'Мазурок І. Є.',
                        'Мазурок Т. Л.',
                        'Ломовцев П. Б.',
                        'Антонова А. Б.',
                        'Селіванова А. В.',
                        'Ольшевская',
                        'Котлик С. В.',
                        'Корнієнко Ю. К.',
                        'Тройнина',
                        'Сіромля С. Г.',
                        'Швець Н. В.',
                        'Попков Д. М.',
                        'Владімірова В. Б.',
                        'Бодюл О. С.']
        self.assisstents.append('Болтач С.В.')
        self.assisstents.append('Асланов')
        self.assisstents.append('Снігур Т.С.')
        self.assisstents.append('Бойцова')
        self.assisstents.append('Кононович')
        self.assisstents.sort()

        self.comboBox_Assisstents.addItems(self.assisstents)
        self.comboBox_Lectors.addItems(self.lectors)
        self.comboBox_Subjects.addItems(self.subjects)
        self.pushButton_plus_asistent.clicked.connect(self.add_assisstant)
        self.pushButton_plus_lector.clicked.connect(self.add_lector)
        self.pushButton_minus_assisstent.clicked.connect(self.minus_assisstant)
        self.pushButton_minus_lector.clicked.connect(self.minus_lector)
        self.pushButton_Calculate_load.clicked.connect(self.calculate_load)

    def calculate_load(self):
        array_lectors = []
        for item in range(self.listWidget_lectors.count()):
            array_lectors.append(self.listWidget_lectors.item(item).text())
        array_assisstents = []
        for item in range(self.listWidget_assisstents.count()):
            array_assisstents.append(self.listWidget_assisstents.item(item).text())
        print(self.comboBox_Subjects.currentText())
        print(array_lectors)
        print(array_assisstents)

    def add_assisstant(self):
        self.listWidget_assisstents.addItem(self.comboBox_Assisstents.currentText())

    def add_lector(self):
        self.listWidget_lectors.addItem(self.comboBox_Lectors.currentText())

    def minus_assisstant(self):
        try:
            self.listWidget_assisstents.takeItem(self.listWidget_assisstents.currentRow())
        except AttributeError:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning,
                                          "Delete item", "You didn't select item", QtWidgets.QMessageBox.Ok)
            error.exec_()

    def minus_lector(self):
        try:
            self.listWidget_lectors.takeItem(self.listWidget_lectors.currentRow())
        except AttributeError:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning,
                                          "Delete item", "You didn't select item", QtWidgets.QMessageBox.Ok)
            error.exec_()

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
                error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning,
                                              "Open Source File", "Failed to read file'", QtWidgets.QMessageBox.Ok)
                error.exec_()
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
        try:
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
        except:
            error = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Report File", "Failed to create report file\n",
                                          QtWidgets.QMessageBox.Ok)
            error.exec_()
        return
