import xlrd
import xlwt
from collections import OrderedDict
import re
from datetime import datetime

workbook = xlrd.open_workbook('123.xls', encoding_override='utf-8')
worksheet = workbook.sheet_by_index(0)
dirs = []


for rownum in range(0,worksheet.nrows):
    row_values = worksheet.row_values(rownum)
    a = 0
    for colnum in row_values:
        if colnum == 'Направление подготовки':
            c = rownum + 1 
            dir = a
        elif colnum == 'Статус заявления': a_status = a
        elif colnum == 'Оригиналы док.-тов': docs = a
        elif colnum == 'Состояние договора': c_status = a
        a+=1

for rownum in range(c, worksheet.nrows):
    counter = OrderedDict()
    row_values = worksheet.row_values(rownum)
    counter['dir'] = row_values[dir]
    counter['docs'] = row_values[docs]
    counter['c_status'] = row_values[c_status]
    counter['a_status'] = row_values[a_status]
    dirs.append(counter)


all_dirs = []

for i in dirs:
    if i['dir'] not in all_dirs:
        all_dirs.append(i['dir'])


font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.bold = True

style0 = xlwt.XFStyle()
style0.font = font0


style1 = xlwt.XFStyle()
style1.name = 'Times New Roman'
style1.num_format_str = 'DD/MM/YY' 


wb = xlwt.Workbook()
ws = wb.add_sheet('List 1')


ws.write(0, 0, 'Результаты за:', style0)
ws.write(0, 1, datetime.now(), style1 )
ws.write(2, 0, 'Направление', style0)
ws.write(2, 1, 'Общ. кол-во \nзаявок', style0)
ws.write(2, 2, 'Оригиналы', style0)
ws.write(2, 3, 'Заключили \nдоговора', style0)
ws.write(2, 4, 'В приказе', style0)


new_row_count = 3
total = total_d = 0
for i in all_dirs:
    count = count_d = count_c = count_a = 0
    for j in dirs:
        if i == j['dir']:
            count += 1
            total += 1
            if j['docs'] == 'Да' : 
                count_d += 1
                total_d += 1
            else: None
            if j['c_status'] == 'Подписан. Внесена оплата' : count_c +=1 
            else: None
            if j['a_status'] == 'В приказе' : count_a += 1
            else: None

    reg = re.compile('\w[..]\.')
    print(reg.sub('', i))
    print(f'Всего заявок: {count}\nИз них:\nПодали Оригиналы: {count_d}\nЗаключили договора: {count_c}\nВ Приказе: {count_a}\n')
    
    
    ws.write(new_row_count, 0, i, style0)
    ws.write(new_row_count, 1, count, style0)
    ws.write(new_row_count, 2, count_d, style0)
    ws.write(new_row_count, 3, count_c, style0)
    ws.write(new_row_count, 4, count_a, style0)
    new_row_count +=1


ws.write(new_row_count, 0, 'Итого:', style0)
ws.write(new_row_count, 1, total, style0)
ws.write(new_row_count, 2, total_d, style0)
wb.save('Результаты.xls')
