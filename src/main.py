import sys
from PyQt5 import QtWidgets, QtGui
from src.QtGUI import AppGUI

if __name__ == "__main__":
    QtWidgets.QApplication.setStyle('Fusion')
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('resources\\логотип_коротко_уменьшено.png'))
    window = AppGUI()
    window.show()
    app.exec_()

