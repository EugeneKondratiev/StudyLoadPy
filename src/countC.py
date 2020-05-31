from openpyxl import *
from openpyxl.worksheet.worksheet import Worksheet

import numpy as np


class CountC:
    def __init__(self, file_name_1sem_den, file_name_1sem_zaoch, file_name_2sem_den, file_name_2sem_zaoch, subject):
        wd = load_workbook(file_name_1sem_den)
        ws = load_workbook(file_name_1sem_zaoch)
        wq = load_workbook(file_name_2sem_den)
        ww = load_workbook(file_name_2sem_zaoch)
        denna2: Worksheet = wq.active
        zaochna2: Worksheet = ww.active
        denna1: Worksheet = wd.active
        zaochna1: Worksheet = ws.active
        predmet = subject
        chasy = []
        chasy2 = []
        lektor = []
        lektor2 = []
        ass = []
        ass2 = []
        den = {}
        zaoch = {}
        c = []

        for i in range(18, denna1.max_row):
            if denna1.cell(row=i, column=2).value != None and denna1.cell(row=i, column=2).value != 0:
                k = denna1.cell(row=i - 1, column=2).value

                dict1_cond = {temp: v for (temp, v) in den.items() if temp == k if temp != None}
                if len(dict1_cond.items()) != 0 and t != "":
                    nympyarray1 = np.array(list(dict1_cond.values()))
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    den[t] = nympyarray3.tolist()
                else:
                    den[k] = c

                c = []
                t = k
                for j in range(16, denna1.max_column + 1):
                    if denna1.cell(row=i, column=j).value is not None and isinstance(denna1.cell(row=i, column=j).value,
                                                                                     str) != True:
                        c.append(denna1.cell(row=i, column=j).value)
        c = []
        t = ""
        y = 0
        for i in range(18, zaochna1.max_row):
            if zaochna1.cell(row=i, column=2).value != None and zaochna1.cell(row=i, column=2).value != "":
                k = zaochna1.cell(row=i - 1, column=2).value
                dict1_cond = {temp: v for (temp, v) in zaoch.items() if temp == k if temp != None}
                if len(dict1_cond.items()) != 0 and t != "":
                    nympyarray1 = np.array(list(dict1_cond.values()))
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    zaoch[t] = nympyarray3.tolist()
                else:
                    zaoch[k] = c
                t = k
                c = []
                for j in range(16, zaochna1.max_column + 1):
                    if zaochna1.cell(row=i, column=j).value is not None and isinstance(
                            zaochna1.cell(row=i, column=j).value, str) != True:
                        c.append(zaochna1.cell(row=i, column=j).value)
        del den[None]
        del zaoch[None]
        dict1_tripleCond = {}
        for (temp, v) in den.items():
            ikval = False
            for (k, c) in zaoch.items():
                if temp == k and temp != None:
                    nympyarray1 = np.array(v)
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    dict1_tripleCond[k] = nympyarray3.tolist()
                    ikval = True
            if ikval == False:
                dict1_tripleCond[temp] = v
                dict1_tripleCond[k] = c

        den2 = {}
        zaoch2 = {}
        c = []
        k = ""
        for i in range(18, denna2.max_row):
            if denna2.cell(row=i, column=2).value != None and denna2.cell(row=i, column=2).value != 0:
                k = denna2.cell(row=i - 1, column=2).value

                dict2_cond = {temp: v for (temp, v) in den.items() if temp == k if temp != None}
                if len(dict2_cond.items()) != 0 and t != "":
                    nympyarray1 = np.array(list(dict2_cond.values()))
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    den2[t] = nympyarray3.tolist()
                else:
                    den2[k] = c

                c = []
                t = k
                for j in range(16, denna2.max_column + 1):
                    if denna2.cell(row=i, column=j).value is not None and isinstance(denna2.cell(row=i, column=j).value,
                                                                                     str) != True:
                        c.append(denna2.cell(row=i, column=j).value)
        c = []
        t = ""
        y = 0
        for i in range(18, zaochna2.max_row):
            if zaochna2.cell(row=i, column=2).value != None and zaochna2.cell(row=i, column=2).value != "":
                k = zaochna2.cell(row=i - 1, column=2).value
                dict2_cond = {temp: v for (temp, v) in zaoch.items() if temp == k if temp != None}
                if len(dict2_cond.items()) != 0 and t != "":
                    nympyarray1 = np.array(list(dict2_cond.values()))
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    zaoch2[t] = nympyarray3.tolist()
                else:
                    zaoch2[k] = c
                t = k
                c = []
                for j in range(16, zaochna2.max_column + 1):
                    if zaochna2.cell(row=i, column=j).value is not None and isinstance(
                            zaochna2.cell(row=i, column=j).value, str) != True:
                        c.append(zaochna2.cell(row=i, column=j).value)

        del den2[None]
        del zaoch2[None]
        dict2_tripleCond = {}
        for (temp, v) in den2.items():
            ikval = False
            for (k, c) in zaoch2.items():
                if temp == k and temp != None:
                    nympyarray1 = np.array(v)
                    nympyarray2 = np.array(c)
                    nympyarray3 = np.array(nympyarray1 + nympyarray2)
                    dict2_tripleCond[k] = nympyarray3.tolist()
                    ikval = True
            if ikval == False:
                dict2_tripleCond[temp] = v
                dict2_tripleCond[k] = c

        for (k, c) in dict1_tripleCond.items():
            if predmet == k:
                chasy = c

        k = ""
        for (y, c) in dict2_tripleCond.items():
            if predmet == y:
                chasy2 = c

        for i in range(0, len(chasy)):
            if i == 0 or i == 3 or i == 4 or i == 5 or i == 6 or i == 7 or i == 8 or i == 9 or i == 10 or i == 11 or i == 12 or i == 15 or i == 16:
                lektor.append(chasy[i])
            elif i != 17:
                ass.append(chasy[i])

        for i in range(0, len(chasy2)):
            if i == 0 or i == 3 or i == 4 or i == 5 or i == 6 or i == 7 or i == 8 or i == 9 or i == 10 or i == 11 or i == 12 or i == 15 or i == 16:
                lektor2.append(chasy2[i])
            elif i != 17:
                ass2.append(chasy2[i])

        # nympyarray1 = np.array(lektor)
        # nympyarray2 = np.array(lektor2)
        # nympyarray3 = np.array(nympyarray1 + nympyarray2)
        # lektor_obch = nympyarray3
        #
        # nympyarray1 = np.array(ass)
        # nympyarray2 = np.array(ass2)
        # nympyarray3 = np.array(nympyarray1 + nympyarray2)
        # ass_obch = nympyarray3

    #CОБСТВЕННО ВОТ ЦиФРА
        self.sem1_ass = np.sum(ass)
        self.sem1_lektor = np.sum(lektor)
        self.sem2_ass = np.sum(ass2)
        self.sem2_lektor = np.sum(lektor2)
