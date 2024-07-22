import openpyxl
import numpy as np

# Nhap du lieu
wb = openpyxl.load_workbook("Data.xlsx")


def importData(sheet):
    return ([[cell.value for cell in row] for row in sheet])


Properties = wb['Properties']
P = importData(Properties['A2:E68'])

for p in P:
    print(p,',')


def inputData():
    # Nhập file dữ liệu.
    Data = wb['Test']
    D = importData(Data)
    return D


def fuzzy(Xi, input1, input2=100):
    if Xi == 18:
        if input1.upper() == "TSG NẶNG":
            return 2
        if input1.upper() == "TSG":
            return 1
        if input1.upper() == "THAI BÌNH THƯỜNG":
            return 0
    if Xi == 2:
        if input1.upper() in P[7][2]:
            return P[7][4]
        elif input1.upper() in P[8][2]:
            return P[8][4]
        else:
            return None
    else:
        for X in P:
            if Xi != X[0]:
                continue
            if input1 is None:
                return None
            else:
                if input2 is None:
                    return None
                else:
                    input = float(input1) / ((float(input2)/100) ** 2)
                    if float(X[2]) <= float(input) <= float(X[3]):
                        return X[4]


def fuzzyList(X):
    list = []
    #print(len(X))
    for i in range(len(X)):
        if i < 4:
            list.append(fuzzy(i, X[i]))
        elif i == 4:
            list.append(fuzzy(i, X[i+1], X[i]))
            i += 1
        elif i > 5:
            list.append(fuzzy(i-1, X[i]))
    return list


def fuzzyData(D):
    newD = []
    for i in range(len(D)):
        X = fuzzyList(D[i])
        newD.append(X)
        #print(X)
    return newD


def writeToExcel(sheet, table, row):
    for x in range(row):
        for y in range(len(table[x])):
            sheet.cell(row=x + 1, column=y + 1, value=table[x][y])


def writeFuzzyDataToExcel():
    D = inputData()
    newD = fuzzyData(D)
    sheet = wb['fuzzyTest']
    writeToExcel(sheet, newD, len(newD))
    wb.save("Data.xlsx")
    # print(newD)


#writeFuzzyDataToExcel()