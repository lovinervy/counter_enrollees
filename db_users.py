import xlrd
import sqlite3 as sql 


class parsing_users():
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.CREATE_DB_USERS()
        self.SEARCH_COLNUM()
        self.PARSING_EXCEL()
        self.EXPORT_DB_ABIT()
        self.FILTERING_DB_ABIT()
        self.INSERT_INTO_DB_USERS()

    
    def CREATE_DB_USERS(self):
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
            print("Done create 'abit'")
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()
    

    def SEARCH_COLNUM(self):
        for rownum in range(0, self.worksheet.nrows):
            row_values = self.worksheet.row_values(rownum)
            for a, colnum in enumerate(row_values):
                if colnum == 'ФИО абитуриента':
                    self.COL_NUM = rownum + 1
                    self.NAME = a   
                elif colnum == 'Номер заявления':
                    self.NUMBER = a
                elif colnum == 'Направление подготовки':
                    self.TRAINING_DIR = a
                elif colnum =='Статус заявления':
                    self.STATUS = a
                elif colnum == 'Средний балл (ЕГЭ)':
                    self.AVERAGE_EGE = a
                elif colnum == 'Оригиналы док.-тов':
                    self.ORIGINAL_DOCS = a
                elif colnum == 'Приоритет в заявлении':
                    self.PRIORITET = a
                elif colnum == 'Особое право':
                    self.SPECIAL = a
                elif colnum == 'Абитуриент-иностранец проверен':
                    self.FOREIGN = a
                elif colnum == 'Нужд. в общ.':
                    self.HOSTEL = a
                elif colnum == 'Тип документа об образовании':
                    self.DOCS = a
                elif colnum == 'Зачислен по направлению':
                    self.ACCESS_DIR = a
                elif colnum == 'Субъект РФ':
                    self.COUNTRY = a
                elif colnum == 'Согласие о зачислении':
                    self.CONSENT = a


    def PARSING_EXCEL(self):
        self.big_data = []
        for rownum in range(self.COL_NUM, self.worksheet.nrows):
            data = []
            row_values = self.worksheet.row_values(rownum)
            data.append(row_values[self.NAME])
            data.append(row_values[self.NUMBER])
            data.append(row_values[self.TRAINING_DIR])
            data.append(row_values[self.STATUS])
            data.append(row_values[self.AVERAGE_EGE])
            data.append(row_values[self.ORIGINAL_DOCS])
            data.append(row_values[self.PRIORITET])
            data.append(row_values[self.SPECIAL])
            data.append(row_values[self.FOREIGN])
            data.append(row_values[self.HOSTEL])
            data.append(row_values[self.DOCS])
            data.append(row_values[self.ACCESS_DIR])
            data.append(row_values[self.COUNTRY])
            data.append(row_values[self.CONSENT])
            self.big_data.append(data)
        print('PARSING EXCEL Done')


    def EXPORT_DB_ABIT(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            cur.execute('SELECT * FROM abit')
            self.result = cur.fetchall()
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()


    def FILTERING_DB_ABIT(self):
        self.FILTERED_ABIT = []
        for bd in self.big_data:
            if tuple(bd) not in self.result:
                if bd not in self.FILTERED_ABIT:
                    self.FILTERED_ABIT.append(bd)

    
    def INSERT_INTO_DB_USERS(self):          
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            for data in self.FILTERED_ABIT:
                cur.execute('INSERT INTO abit VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', data)
            con.commit()        
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()
