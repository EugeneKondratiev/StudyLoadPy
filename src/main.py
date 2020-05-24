from src.DennaSem1 import DennaSem1
from src.DennaSem2 import DennaSem2
from src.ZaochnaSem2 import ZaoсhnaSem2
from src.ZaoсhnaSem1 import ZaoсhnaSem1
from src.pars import pars
from src.tGUI import StudyLoadFrame

p1 = DennaSem1()
p2 = ZaoсhnaSem1()
p4 = DennaSem2()
p3 = ZaoсhnaSem2()
p5 = pars()
if __name__ == "__main__":
    StudyLoadFrame().mainloop()

