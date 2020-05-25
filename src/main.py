import sys
from PyQt5 import QtWidgets, QtGui

from src.DennaSem1 import DennaSem1
from src.DennaSem2 import DennaSem2
from src.QtGUI import AppGUI
from src.ZaochnaSem2 import ZaoсhnaSem2
from src.ZaoсhnaSem1 import ZaoсhnaSem1

#p1 = DennaSem1()
#p2 = ZaoсhnaSem1()
#p4 = DennaSem2()
#p3 = ZaoсhnaSem2()
#p5 = pars()
if __name__ == "__main__":
    QtWidgets.QApplication.setStyle('Fusion')
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('resources\\логотип_коротко_уменьшено.png'))
    window = AppGUI()
    window.show()
    app.exec_()


