import sqlite3 as sql
import sys
import xlrd


con = None
def CREATE_DATABASE():
    try:
        con = sql.connect('DATABASE.db')
        cur = con.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS abit(ФИО TEXT,' 
                                                    'Номер_заявления TEXT, '
                                                    'Направление_подготвки TEXT, '
                                                    'Статус_заявления TEXT, '
                                                    'Средний_балл_ЕГЭ FLOAT, '
                                                    'Оригинал_документов TEXT, '
                                                    'Приоритет TEXT, '
                                                    'Особое_право TEXT, '
                                                    'Иностранец TEXT, '
                                                    'Нуждается_в_общежитии TEXT, '
                                                    'Тип_документа TEXT, '
                                                    'Зачислен_по_направлению TEXT, '
                                                    'Субъект_РФ TEXT, '
                                                    'Согласие_на_зачисление TEXT)')
        print("Done create database")
        
    
    except sql.Error as s:
        print(f'Error: {s}')

    finally:
        if con:
            cur.close()
            con.close()


def INSERT_INTO_DATABASE(big_data):              
    try:
        con = sql.connect('DATABASE.db')
        cur = con.cursor()
        for data in big_data:
            cur.execute('INSERT INTO abit VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', data)
        con.commit()
        
    
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
                print('FIO Done')
            elif colnum == 'Номер заявления':
                NUMBER = a
                print('NUMBER Done')
            elif colnum == 'Направление подготовки':
                TRAINING_DIR = a
                print('T_DIR Done')
            elif colnum =='Статус заявления':
                STATUS = a
                print('STATUS Done')
            elif colnum == 'Средний балл (ЕГЭ)':
                AVERAGE_EGE = a
                print('A_EGE Done')
            elif colnum == 'Оригиналы док.-тов':
                ORIGINAL_DOCS = a
                print('O_DOCS Done')
            elif colnum == 'Приоритет в заявлении':
                PRIORITET = a
                print('PRIORITET Done')
            elif colnum == 'Особое право':
                SPECIAL = a
                print('SPECIAL Done')
            elif colnum == 'Абитуриент-иностранец проверен':
                FOREIGN = a
                print('FOREIGN Done')
            elif colnum == 'Нужд. в общ.':
                HOSTEL = a
                print('HOSTEL Done')
            elif colnum == 'Тип документа об образовании':
                DOCS = a
                print('DOCS Done')
            elif colnum == 'Зачислен по направлению':
                ACCESS_DIR = a
                print('A_DIR Done')
            elif colnum == 'Субъект РФ':
                COUNTRY = a
                print('COuNTRY Done')
            elif colnum == 'Согласие о зачислении':
                CONSENT = a
                print('CONSENT Done')
    return {'COL_NUM': COL_NUM, 'NAME': NAME,'NUMBER': NUMBER, 'TRAINING_DIR': TRAINING_DIR, 'STATUS': STATUS, 'AVERAGE_EGE':AVERAGE_EGE,\
            'ORIGINAL_DOCS': ORIGINAL_DOCS, 'PRIORITET': PRIORITET, 'SPECIAL': SPECIAL, 'FOREIGN': FOREIGN,\
            'HOSTEL': HOSTEL, 'DOCS': DOCS, 'ACCESS_DIR': ACCESS_DIR, 'COUNTRY': COUNTRY, 'CONSENT': CONSENT}


def PARSING_EXCEL(hf, ws):
    big_data = []
    for rownum in range(hf['COL_NUM'], ws.nrows):
        data = []
        row_values = ws.row_values(rownum)
        data.append(row_values[hf['NAME']])
        data.append(row_values[hf['NUMBER']])
        data.append(row_values[hf['TRAINING_DIR']])
        data.append(row_values[hf['STATUS']])
        data.append(row_values[hf['AVERAGE_EGE']])
        data.append(row_values[hf['ORIGINAL_DOCS']])
        data.append(row_values[hf['PRIORITET']])
        data.append(row_values[hf['SPECIAL']])
        data.append(row_values[hf['FOREIGN']])
        data.append(row_values[hf['HOSTEL']])
        data.append(row_values[hf['DOCS']])
        data.append(row_values[hf['ACCESS_DIR']])
        data.append(row_values[hf['COUNTRY']])
        data.append(row_values[hf['CONSENT']])
        big_data.append(data)
    print('PARSING EXCEL Done')
    return big_data

'''
def WRITE_TO_DATABASE(data):
    for i in data:'''
        





if __name__ == "__main__":
    workbook = xlrd.open_workbook('123.xls', encoding_override='utf-8')
    worksheet = workbook.sheet_by_index(0)
    hash_file = SEARCH_COLNUM(worksheet)
    data = PARSING_EXCEL(hf=hash_file, ws=worksheet)
    CREATE_DATABASE()
    INSERT_INTO_DATABASE(data)