#README
"""
усі значення для образунків беруться з файлу example.xlsx
який знаходиться у папці з програмою
будь які значення в файлі можна змінювати
"""

#імпорт бібліотек
import openpyxl
import sys

#завантаження таблиці
wb = openpyxl.load_workbook('example.xlsx')

#вибір листа
shs = wb.sheetnames
sh = wb[shs[0]]

#перевірки сум коефіцієнтів
chek_l = round(sh['A2'].value + sh['A3'].value + sh['A4'].value + sh['A5'].value + sh['A6'].value)
chek_s = round(sh['C2'].value + sh['D3'].value + sh['E4'].value)
chek_p = round(sh['C3'].value + sh['D3'].value + sh['E3'].value)
chek_b = round(sh['C4'].value + sh['D4'].value + sh['E4'].value)
chek_d = round(sh['C5'].value + sh['D5'].value + sh['E5'].value)
chek_c = round(sh['C6'].value + sh['D6'].value + sh['E6'].value)

if chek_l != 1 or chek_s != 1 or chek_p != 1 or chek_b != 1 or chek_d != 1 or chek_c != 1:
    print("error!")
    sys.exit()

#підрахунки
print("success!")
res_i = round(sh['C2'].value * sh['A2'].value + sh['C3'].value * sh['A3'].value + sh['C4'].value * sh['A4'].value + sh['C5'].value * sh['A5'].value + sh['C6'].value * sh['A6'].value, 4)
res_x = round(sh['D2'].value * sh['A2'].value + sh['D3'].value * sh['A3'].value + sh['D4'].value * sh['A4'].value + sh['D5'].value * sh['A5'].value + sh['D6'].value * sh['A6'].value, 4)
res_s = round(sh['E2'].value * sh['A2'].value + sh['E3'].value * sh['A3'].value + sh['E4'].value * sh['A4'].value + sh['E5'].value * sh['A5'].value + sh['E6'].value * sh['A6'].value, 4)

#виведення результатів та відповіді
print(sh['C1'].value,"-",res_i)
print(sh['D1'].value,"-",res_x)
print(sh['E1'].value,"-",res_s)

if res_i > res_x and res_i > res_s:
    print("відповідь: ",sh['C1'].value)
elif res_x > res_i and res_x > res_s:
    print("відповідь: ",sh['D1'].value)
elif res_s > res_i and res_s > res_x:
    print("відповідь:",sh['E1'].value)

wb.save('example.xlsx')
