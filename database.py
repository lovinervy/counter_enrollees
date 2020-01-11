import sqlite3 as sql
import sys
import xlrd


con = None
def CREATE_DATABASE():
    try:
        con = sql.connect('DATABASE.db')
        cur = con.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS ABIT(ФИО TEXT,' 
                                                    'Направление_подготвки TEXT, '
                                                    'Статус_заявления TEXT, '
                                                    'Средний_балл_ЕГЭ FLOAT, '
                                                    'Оригинал_документов FLOAT, '
                                                    'Приоритет INT, '
                                                    'Особое_право TEXT, '
                                                    'Иностранец INT, '
                                                    'Нуждается_в_общежитии INT, '
                                                    'Тип_документа TEXT, '
                                                    'Зачислен_по_направлению INT, '
                                                    'Субъект_РФ TEXT, '
                                                    'Согласие_на_зачисление INT)')
        print("Done")
    
    except sql.Error as s:
        print(f'Error: {s}')

    finally:
        if con:
            cur.close()
            con.close()


def IMPORT_TO_DATABASE(big_data):               #Пофикси это дерьмо, под фиксом имею ввиду переделай
    try:
        con.sql.connect('DATABASE.db')
        cur = con.cursor()
        for data in big_data:
            cur.execute(data)
    
    except sql.Error as s:
        print(f'Error: {s}')
    
    finally:
        cur.close()
        con.close()
    
def SEARCH_COLNUM(ws):
    for rownum in range(0, ws.nrows):
        row_values = ws.row_values(rownum)
        for a, colnum in enumerate(row_values):
            if colnum == 'ФИО абитуриента':
                COL_NUM = rownum + 1
                NAME = a
            elif colnum == 'Направление подготовки':
                TRAINING_DIR = a
            elif colnum =='Статус заявления':
                STATUS = a
            elif colnum == 'Средний балл (ЕГЭ)':
                AVERAGE_EGE = a
            elif colnum == 'Оригиналы док.-тов':
                ORIGINAL_DOCS = a
            elif colnum == 'Приоритет в заявлении':
                PRIORITET = a
            elif colnum == 'Особое право':
                SPECIAL = a
            elif colnum == 'Абитуриент-иностранец проверен':
                FOREIGN = a
            elif colnum == 'Нужд. в общ.':
                HOSTEL = a
            elif colnum == 'Тип документа об образовании':
                DOCS = a
            elif colnum == 'Зачислен по направлению':
                ACCESS_DIR = a
            elif colnum == 'Субъект РФ':
                COUNRTY = a
            elif colnum == 'Согласие о зачислении':
                CONSENT = a
    return {'COL_NUM': COL_NUM, 'NAME': NAME, 'TRAINING_DIR': TRAINING_DIR, 'STATUS': STATUS, 'AVERAGE_EGE':AVERAGE_EGE,\
            'ORIGINAL_DOCS': ORIGINAL_DOCS, 'PRIORITET': PRIORITET, 'SPECIAL': SPECIAL, 'FOREIGN': FOREIGN,\
            'HOSTEL': HOSTEL, 'DOCS': DOCS, 'ACCESS_DIR': ACCESS_DIR, 'COUNRTY': COUNRTY, 'CONSENT': CONSENT}


def PARSING_EXCEL(hf, ws):
    big_data = []
    for rownum in range(hf['COL_NUM'], ws.nrows):
        data = []
        row_values = ws.row_values(rownum)
        data.append(row_values[hf['NAME']])
        data.append(row_values[hf['TRAINING_DIR']])
        data.append(row_values[hf['STATUS']])
        data.append(row_values[hf['AVERAGE_EGE']])
        data.append(row_values[hf['ORIGINAL_DOCS']])
        data.append(row_values[hf['PRIORITET']])
        data.append(row_values[hf['SPECIAL']])
        data.append(row_values[hf['FOREIGN']])
        data.append(row_values[hf['HOSTEL']])
        data.append(row_values[hf['DOCS']])
        data.append(row_values[hf['ACCES_DIR']])
        data.append(row_values[hf['COUNTRY']])
        data.append(row_values[hf['CONSENT']])
        big_data.append(data)
    return big_data


def WRITE_TO_DATABASE(data):
    for i in data:
        





if __name__ == "__main__":
    workbook = xlrd.open_workbook('123.xls', encoding_override='utf-8')
    worksheet = workbook.sheet_by_index(0)

    

CREATE_DATABASE()