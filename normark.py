import xlrd, csv

wb = xlrd.open_workbook('SpbFishZakaz.xlsx')

sheet = wb.sheet_by_name('Ост')
lst_ref = []
for row in range(sheet.nrows):
    ref = sheet.cell_value(row, 0)
    price = str(sheet.cell_value(row, 3)).replace('В наличии', '777').replace('Меньше 3 шт', '2')
    if ref != '':
        lst_ref.append([ref, price])

with open('Ballance_NORMARK.csv', 'w', newline='') as csv_r:
    csv_w = csv.writer(csv_r, delimiter=',')
    csv_w.writerows(lst_ref)