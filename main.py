import xlrd, csv

wb = xlrd.open_workbook('Минотавр СПб прайс01.02.xlsx')

sheet1 = wb.sheet_by_index(0)
lst_balanc = []
for row in range(sheet1.nrows):
    artik = str(sheet1.cell_value(row, 1)).strip()
    balanc = str(sheet1.cell_value(row, 3)).strip().replace('есть', '777')
    artik_balanc = (artik + ';' + balanc).splitlines()
    price_col = str(sheet1.cell_value(row, 7)).replace('Цена клиента', '')
    price_col2 = price_col
    # csv_pr = (artik + ';' + price_col2)
    print(price_col2)
    for p in price_col2:
        if len(p.strip()) != 0 and p.strip() != '.':
            p = float(p)
            # if p < 200:
            #     p *= 1.78
    #         elif 200<= float(p) <400:
    #             p = p * 1,55
    #         elif 400<= float(p) <900:
    #             p = p * 1,47
    #         elif 900<= float(p) <2000:
    #             p = p * 1,43
    #         else: p = p * 1.38


    # print(price_col2)



#     for a_b in artik_balanc:
#         if len(a_b) > 1:
#             lst_balanc.append(a_b)
# csv_balance = lst_balanc[3::]
#
# with open('MINO_Ballance.csv', 'w', newline='') as csv_f:
#     csv_w = csv.writer(csv_f, delimiter = ',')
#     for item in csv_balance:
#         csv_w.writerow([item])