import xlrd
import sqlite3 as sql


class parsing_training_dir():
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.CREATE_DB_TRAINING_DIR()
        self.SEARCH_COLNUM()
        self.PARSING_EXCEL()
        self.EXPORT_DB_TRAINING_DIR()
        self.FILTERING_DB_TRAINING_DIR()
        self.INSERT_INTO_DB_TRAININF_DIR()


    def CREATE_DB_TRAINING_DIR(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            cur.execute('CREATE TABLE IF NOT EXISTS training_dir(Номер_заявление TEXT,'
                                                                'Направление TEXT,'
                                                                'Приоритет TEXT,'
                                                                'Зачислен_по_направлению TEXT)')
            print("Done create 'training dir'")
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()


    def SEARCH_COLNUM(self):
        for rownum in range(0, self.worksheet.nrows):
            row_values = self.worksheet.row_values(rownum)
            for num, colnum in enumerate(row_values):
                if colnum == 'Номер заявления':
                    self.COL_NUM = rownum + 1
                    self.NUMBER = num
                elif colnum == 'Направление подготовки':
                    self.TRAINING_DIR = num
                elif colnum == 'Приоритет в заявлении':
                    self.PRIORITET = num
                elif colnum == 'Зачислен по направлению':
                    self.ACCESS_DIR = num
        print('SEARCH CONUM Done')

    
    def PARSING_EXCEL(self):
        self.big_data = []
        for rownum in range(self.COL_NUM, self.worksheet.nrows):
            data = []
            row_values = self.worksheet.row_values(rownum)
            data.append(row_values[self.NUMBER])
            data.append(row_values[self.TRAINING_DIR])
            data.append(row_values[self.PRIORITET])
            data.append(row_values[self.ACCESS_DIR])
            self.big_data.append(data)
        print('PARSING EXCEL Done')


    def EXPORT_DB_TRAINING_DIR(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            cur.execute('SELECT * FROM training_dir')
            self.result = cur.fetchall()
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()
                print('Export_db Done')


    def FILTERING_DB_TRAINING_DIR(self):
        self.FILTERING_DIR = []
        for db in self.big_data:
            if tuple(db) not in self.result:
                if db not in self.FILTERING_DIR:
                    self.FILTERING_DIR.append(db)


    def INSERT_INTO_DB_TRAININF_DIR(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            for data in self.FILTERING_DIR:
                cur.execute('INSERT INTO training_dir VALUES(?, ?, ?, ?)', data)
            con.commit()
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()
                print('All Done')