import win32com.client
import sys
import decimal
from collections import namedtuple

nameSheets = ('Результаты', 'Результаты2')


def addRow(table, nodenumber, support, x, y, z, xo, yo, zo):
    l = [nodenumber, support, decimal.Decimal(x), decimal.Decimal(y), decimal.Decimal(z), decimal.Decimal(xo),
         decimal.Decimal(yo), decimal.Decimal(zo)]
    table.append(l)


def mergeTable(table1: list, table2: list):
    result = []
    lNode = [n[0] for n in table2]
    for row in table1:
        try:
            i = lNode.index(row[0])
            row.append(table2.pop(i))
            lNode.pop(i)
        except:
            row.append(None)
    for row in table2:
        table1.append([row[0], row[1], *[None for i in range(6)], row])


if len(sys.argv) < 2:
    raise ('Не указаны все аргументы! "имя файла 1" "имя файла 2"')

if sys.argv[0].find('/') >= 0:
    sep = '/'
else:
    sep = '\\'
curdir = sys.argv[0][:sys.argv[0].rfind(sep[0]) + 1]
fileName1 = sys.argv[1]
fileName2 = sys.argv[2]

exApp = win32com.client.Dispatch("Excel.Application")
wb = exApp.Workbooks.Open(curdir + fileName1)

for sheet in wb.Sheets:
    if sheet.Name in nameSheets:
        # sheet = wb.Sheets("Результаты")
        sheet.Delete()

sheet = wb.ActiveSheet

table1 = []
i_header = 0
for (row) in sheet.UsedRange.Rows:
    lrow = []
    for cell in row.Cells:
        lrow.append(cell.Text)
    if i_header > 1:
        addRow(table1, *lrow[:-1])
    i_header += 1

wb2 = exApp.Workbooks.Open(curdir + fileName2)

sheet2 = wb2.ActiveSheet
table2 = []
i_header = 0
for (row) in sheet2.UsedRange.Rows:
    lrow = []
    for cell in row.Cells:
        lrow.append(cell.Text)
    if i_header > 1:
        addRow(table2, *lrow[:-1])
    i_header += 1
wb2.Close()

mergeTable(table1, table2)

newSheet = wb.Sheets.Add()
newSheet.Name = nameSheets[0]
newSheet.Activate()
# newSheet.Cells.NumberFormat = '@'
i = 3
for row in table1:
    j = 1
    for cell in row:
        if type(cell) is list:
            k = 0
            for item in cell:
                if k > 1:
                    val = item
                    newSheet.Cells(i, j + k).Value = val
                    newSheet.Cells(i, j + k).NumberFormat = '0.00'
                k += 1
        else:
            newSheet.Cells(i, j).Value = cell
            if j < 3:
                newSheet.Cells(i, j).NumberFormat = '@'
            else:
                newSheet.Cells(i, j).NumberFormat = '0.00'
        j += 1
    i += 1

newSheet = wb.Sheets.Add()
newSheet.Name = nameSheets[1]
newSheet.Activate()

tableNewLine = [row[8] for row in table1 if row[8] and row[2] == None]
tableDeleteLine = [row[:-1] for row in table1 if not row[8]]
tableDifference = []
for row in table1:
    if (type(row[8]) is list) and row[2]:
        tableDifference.append([row[0], row[1],
                                row[2] - row[8][2],
                                row[3] - row[8][3],
                                row[4] - row[8][4],
                                row[5] - row[8][5],
                                row[6] - row[8][6]])
i = 3
for row in tableDifference:
    sum = 0
    for cell in row[2:]:
        sum += cell
    if sum == 0:
        continue
    j = 1
    for cell in row:
        if j > 2:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '0.00'
        else:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '@'
        j += 1
    i += 1

i += 1
newSheet.Cells(i, 1).Value = 'Удаленные узлы'
i += 1
for row in tableDeleteLine:
    j = 1
    for cell in row:
        if j > 2:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '0.00'
        else:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '@'
        j += 1
    i += 1

i += 1
newSheet.Cells(i, 1).Value = 'Новые узлы'
i += 1
for row in tableNewLine:
    j = 1
    for cell in row:
        if j > 1:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '0.00'
        else:
            newSheet.Cells(i, j).Value = cell
            newSheet.Cells(i, j).NumberFormat = '@'
        j += 1
    i += 1

wb.Save()
wb.Close()
exApp.Quit()
