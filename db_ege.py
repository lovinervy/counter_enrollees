import xlrd
import sqlite3 as sql


class parsing_exams():
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.CREATE_DB_EXAMS()
        self.EXPORT_DB_EXAMS()
        self.FIND_EXAMS()
        self.COMPLETE_USER_EXAMS()
        self.CREATE_ID()
        self.PARSING_EXAM()
        self.FILTERING_EXAMS()
        self.INSERT_INTO_DB_EXAMS()


    def CREATE_DB_EXAMS(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            cur.execute('CREATE TABLE IF NOT EXISTS exams(КОД INT,'
                                                        'ЭКЗАМЕН TEXT,'
                                                        'ТИП_ЭКЗАМЕНА TEXT,'
                                                        'БАЛЛ INT)')
            print("Done create exams")       
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()


    def EXPORT_DB_EXAMS(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            cur.execute('SELECT * FROM exams')
            self.result = cur.fetchall()
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()


    def FIND_EXAMS(self):
        for i in range(0, self.worksheet.nrows):
            row_values = self.worksheet.row_values(i)
            for c, j in enumerate(row_values):
                if j == 'Номер заявления':
                    self.count_id = c
                elif j == 'Баллы':
                    self.start_row = i + 1
                    self.count_ege = c
                elif j == 'Предметы (ЕГЭ)':
                    self.count_dir_ege = c

    
    def COMPLETE_USER_EXAMS(self):
        EXAMS = []
        for rownum in range(self.start_row, self.worksheet.nrows):
            USER_EXAMS = []
            row_values = self.worksheet.row_values(rownum)
            USER_EXAMS.append(row_values[self.count_id])
            USER_EXAMS.append(row_values[self.count_dir_ege])
            USER_EXAMS.append(row_values[self.count_ege])
            EXAMS.append(USER_EXAMS)
        self.ALL_EXAMS = EXAMS

    
    def CREATE_ID(self):
        for num,user in enumerate(self.ALL_EXAMS):
            user_id = user[0].split('-')
            if user_id[0] == '225':
                user_id = int('1' + user_id[1])
                self.ALL_EXAMS[num].pop(0)
                self.ALL_EXAMS[num].insert(0, user_id)
            elif user_id[0] == '226':
                user_id = int('2' + user_id[1])
                self.ALL_EXAMS[num].pop(0)
                self.ALL_EXAMS[num].insert(0, user_id)
            elif user_id[0] == '227':
                user_id = int('3' + user_id[1])
                self.ALL_EXAMS[num].pop(0)
                self.ALL_EXAMS[num].insert(0, user_id)          
    

    def PARSING_EXAM(self):
        FINAL_MAS = []
        for num, user in enumerate(self.ALL_EXAMS):
            user_exams = user[1].split('\n')
            if user_exams[-1] == '':
                user_exams.pop(-1)
            FOUND_TYPE_EXAM = []
            for finding_type_exam in user_exams:
                type_exam = finding_type_exam.split(': ')
                FOUND_TYPE_EXAM.append(type_exam)
            user_exam_points = user[2].split('\n')
            if user_exam_points[-1]=='':
                user_exam_points.pop(-1)
            for enum, exam in enumerate(user_exams):
                MAS = []
                MAS.append(self.ALL_EXAMS[num][0])
                MAS.append(FOUND_TYPE_EXAM[enum][0])
                try:
                    MAS.append(FOUND_TYPE_EXAM[enum][1])
                except:
                    MAS.append('ИД')
                try:
                    MAS.append(int(user_exam_points[enum]))
                except:
                    MAS.append(0)
                
                FINAL_MAS.append(MAS)
        self.COMPLETED_EXAM = FINAL_MAS
    

    def FILTERING_EXAMS(self):
        self.FILTERED_EXAMS = []
        for bd in self.COMPLETED_EXAM:
            if tuple(bd) not in self.result:
                if bd not in self.FILTERED_EXAMS:
                    self.FILTERED_EXAMS.append(bd)

    
    def INSERT_INTO_DB_EXAMS(self):
        try:
            con = sql.connect('DATABASE.db')
            cur = con.cursor()
            for exam in self.FILTERED_EXAMS:
                cur.execute('INSERT INTO exams VALUES(?, ?, ?, ?)', exam)
            con.commit()
        except sql.Error as s:
            print(f'Error: {s}')
        finally:
            if con:
                cur.close()
                con.close()
