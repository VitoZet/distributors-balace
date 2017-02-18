import xlrd

wb = xlrd.open_workbook('TestMino.xlsx')

sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)

# print(sheet1.cell_value(0, 0)) #выводим первую ячейку
# print(sheet1.nrows) #Считаем кол-во строк на листе
# print(sheet2.nrows)
# print(sheet1.ncols) #Считаем столбцы

# for col in range(sheet1.ncols): #генерим строки
#     print(sheet1.cell_value(5 ,col))

for row in range(sheet1.nrows):  # генерим столбцы
    x = str(sheet1.cell_value(row, 1))
    y = str(sheet1.cell_value(row, 3))
    for yy in y:
        yy.replace('есть', '777')
        print(type(yy))
    balance = [x + ',' + y]
    # print(balance)
    # for b in balance:
    #     b.replace('есть', '777')
    #     print(b)
    # # for b in balance:
    #     b.replace('есть','777')
    #     print(b)
    # # balance = str(sheet1.cell_value(row, 1)) + ',' + str(sheet1.cell_value(row, 3))
# data = [[sheet1.cell_value(r,c) for c in range(sheet1.ncols)] for r in range(sheet1.nrows)] #Лист ексель В лист
