from enum import Flag
import random
from openpyxl.styles import Border, Side, Alignment
import openpyxl as op

def create_excel(r, name):
    excelfile = op.Workbook()
    ws = excelfile['Sheet']
    ws.column_dimensions['A'].width =10
    ws.column_dimensions['B'].width =10
    ws.column_dimensions['C'].width =10
    ws.column_dimensions['D'].width =10
    ws.column_dimensions['E'].width =10
    ws.column_dimensions['F'].width =10
    ws.column_dimensions['G'].width =10
    ws.column_dimensions['H'].width =10
    ws.column_dimensions['I'].width =10
    ws.row_dimensions[1].height =20
    ws.row_dimensions[2].height =20
    ws.row_dimensions[3].height =20
    ws.row_dimensions[4].height =20
    ws.row_dimensions[5].height =20
    ws.row_dimensions[6].height =20
    ws.row_dimensions[7].height =20
    ws.row_dimensions[8].height =20
    ws.row_dimensions[9].height =20
    write = excelfile.active

    row = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    side1 = Side(style='thin', color='000000')
    border_aro = Border(top=side1, bottom=side1, left=side1, right=side1)
    for y in range(0, 9):
        for x in range(0, 9):
            write[f'{row[y]}{x + 1}'].border = border_aro
            write[f'{row[y]}{x + 1}'].alignment = Alignment(horizontal = 'center', vertical = 'center')
            write[f'{row[y]}{x + 1}'].font = op.styles.fonts.Font(size=20)
            if r[y][x] == 0:
                write[f'{row[y]}{x + 1}'].value = ''
            else:
                write[f'{row[y]}{x + 1}'].value = r[y][x]

    excelfile.save(name)

def create():
    r = []
    while len(r) < 9:
        num = random.randint(1, 9)
        if not num in r:
            r.append(num)
    
    return r

def create_row(r, value, rang):
    while 1:
        flag = False
        a = create()
        for x in range(0, 9):
            for y in range(0, 9):
                if r[y][x] == a[x]:
                    flag = True
        if flag == True:
            continue

        for i in range(0, 3):
            for z in range(0, 3):
                for row in range(rang, rang + 3):
                    if r[row][z] == a[i]:
                        flag = True
        for i in range(3, 6):
            for z in range(3, 6):
                for row in range(rang, rang + 3):
                    if r[row][z] == a[i]:
                        flag = True
        for i in range(6, 9):
            for z in range(6, 9):
                for row in range(rang, rang + 3):
                    if r[row][z] == a[i]:
                        flag = True
        if flag == True:
            continue

        for z in range(0, 9):
            r[value][z] = a[z]
        break

    return r

def main():
    r = [[0 for x in range(0, 9)] for y in range(0, 9)]
    for x in range(0, 3):
        r = create_row(r, x, 0)
    for x in range(3, 6):
        r = create_row(r, x, 3)
    for x in range(6, 9):
        r = create_row(r, x, 6)

    create_excel(r, 'solve.xlsx')

    for x in range(0, 20):
        x = random.randint(0, 8)
        y = random.randint(0, 8)
        r[y][x] = 0
    
    create_excel(r, 'problem.xlsx')

if __name__ == '__main__':
    main()