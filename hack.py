import openpyxl
import time
import os
from collections import Counter
from openpyxl.styles import Font, Color, Border, colors, Side

start_time = time.time()

print('File name (example - "ot4et")')
name = str(input('File name : ')) + '.xlsx'

path = os.path.dirname(os.path.realpath(__file__)) + "\\" + name

book = openpyxl.open(path)
sheet = book.active

def clear(l1, l2):
    global updated_b, updated_a
    c1 = Counter(l1)
    c2 = Counter(l2)
    diff = c1-c2
    diff2 = c2-c1
    updated_a = list(diff.elements())
    updated_b = list(diff2.elements())

def write(name, array, col):
    for i in range(1, len(sheet[name])+1):
        if sheet[i][col].value == None:
            sheet[i][col].value = 0
        else:
            if sheet[i][col].value < 0:
                sheet[i][col].value *= -1
            array.append(sheet[i][col].value)
    array.sort()

def update_excel(arr, column, name="Column"):
    thin = Side(border_style="thin", color="000000")
    brdr = Border(left=thin, right=thin, bottom=thin)
    sheet[column][0].value = name
    sheet[column][0].font = Font(color="FF0000", bold=True)
    sheet[column][0].border = Border(left=thin, right=thin, bottom=Side(border_style="double", color="000000"))
    for i in range(len(arr)):
        sheet[column][i+1].value = arr[i]
        sheet[column][i+1].border = brdr

a_column = []
b_column = []
write("A", a_column, 0)
write("B", b_column, 1)

clear(a_column, b_column)

update_excel(updated_a, "D", "A Column")
update_excel(updated_b, "E", "B Column")

print("Updated!")
book.save(path) # or name
print("--- %s seconds ---" % round((time.time() - start_time), 2))