import xlrd, csv

wb = xlrd.open_workbook('Минотавр СПб прайс01.02.xlsx')

sheet1 = wb.sheet_by_index(0)
csv_price = []
csv_balance = []
for row in range(sheet1.nrows):
    try:
        row_num = float(str(sheet1.cell_value(row, 0)).strip())
    except:
        row_num = 0
    if row_num > 0:
        artik = str(sheet1.cell_value(row, 1)).strip()
        balanc = int(float(str(sheet1.cell_value(row, 3)).strip().replace('есть', '777')))
        price_col = float(str(sheet1.cell_value(row, 7)).replace('Цена клиента', ''))
        if price_col < 200:
            price_col *= 1.78
        elif 200 <= price_col < 400:
            price_col *= 1.55
        elif 400 <= price_col < 900:
            price_col *= 1.47
        elif 900 <= price_col < 2000:
            price_col *= 1.43
        else: price_col *= 1.38
        price_col_round = int(round(price_col,-1))


        csv_price.append([artik, price_col_round])
        csv_balance.append([artik,balanc])

with open('Ballance_MINO.csv', 'w', newline='') as csv_f:
    csv_w = csv.writer(csv_f, delimiter = ',')
    csv_w.writerows(csv_balance)

with open('Price_MINO.csv', 'w', newline='') as csv_p:
    csv_wp = csv.writer(csv_p, delimiter = ',')
    csv_wp.writerows(csv_price)