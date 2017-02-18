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
    x = str(sheet1.cell_value(row, 1)).strip()
    y = str(sheet1.cell_value(row, 3)).strip().replace('есть', '777')
    balance = x + ',' + y
    if len(balance)==1:
        
    print(len(balance))
    # if len(balance)<2 in balance:
    #     balance.remove(',')
    # print(balance)
    # for b in balance:
    #     b.replace('есть', '777')
    #     print(b)
    # # for b in balance:
    #     b.replace('есть','777')
    #     print(b)
    # # balance = str(sheet1.cell_value(row, 1)) + ',' + str(sheet1.cell_value(row, 3))
# data = [[sheet1.cell_value(r,c) for c in range(sheet1.ncols)] for r in range(sheet1.nrows)] #Лист ексель В лист
