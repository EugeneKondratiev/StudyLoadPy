import sys
from PyQt5 import QtWidgets, QtGui


from src.navantazhenya import Ui_MainWindow


class ExampleApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def main():
    QtWidgets.QApplication.setStyle('Fusion')
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('resources\\логотип_коротко_уменьшено.png'))
    window = ExampleApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
