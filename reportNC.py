import openpyxl
from openpyxl import load_workbook
from datetime import datetime

wb = load_workbook(filename = "caw_wb.xlsx", data_only=True)        # find NC spreadsheet
sheet = wb._sheets[0]       # apply focus to worksheet

def load_Title():       # load title date to title storage array

    titleLine = []
    colTrace = 1        # counter to trace over columns
    while sheet.cell(2, colTrace).value:        # while column specified has a value, run code
        cellVal = sheet.cell(2, colTrace).value     # instantiating current cell as cellVal
        titleLine.append(cellVal)       # adding to titleLine
        colTrace += 1       # incrementing counter
    return titleLine        # return storage array

titleLine = load_Title()
dueDatePos = titleLine.index('DATE DUE') + 1
statusPos = titleLine.index('DATE COMPLETED') + 1
dueCol = sheet.cell(1, dueDatePos).column_letter

def load_due_dates():
    dueDates = []
    rowTrace = 3
    x = 1
    while sheet['{x}'.format(x=dueCol) + str(rowTrace)].value:
        dueVal = sheet['{x}'.format(x=dueCol) + str(rowTrace)].value
        dueDates.append([x, dueVal])
        rowTrace += 1
        x += 1
    for x in dueDates:
        if len(str(x[1])) > 10:
            x[1] = x[1].strftime('%d/%m/%Y')
    print(dueDates)

load_due_dates()

today = datetime.now().strftime('%d/%m/%Y')


