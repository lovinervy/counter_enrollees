import xlrd
import xlwt
import re
import os

class counter:
    def __init__(self):
        PATH = ''
        files = os.listdir()
        for f in files:
            if f.endswith('.xls'):
                data = f
                break
    
        workbook = xlrd.open_workbook(data, encoding_override='utf-8')
        self.worksheet = workbook.sheet_by_index(0)
        self.find_direction()
        self.take_all_dirs()
        self.delete_elements()
        self.dirs.sort()
        self.create_direction()
        self.count()
    

    def find_direction(self):
        for rownum in range(0, self.worksheet.nrows):
            row_values = self.worksheet.row_values(rownum)
            for num, col in enumerate(row_values):
                if col == 'Конкурсные группы':
                    self.col = rownum + 1
                    self.row = num
                elif col == 'Статус заявления':
                    self.row_status = num


    def take_all_dirs(self):
        dirs = []
        for rownum in range(self.col, self.worksheet.nrows):
            row_values = self.worksheet.row_values(rownum)
            if row_values[self.row_status] in ['Новое', 'Принято']:
                dirs.append(row_values[self.row])
        self.dirs = dirs
    
    def delete_elements(self):
        dirs = []
        del_elements = ['/', 'ориг.', 'бак.', 'бюдж.', 'дог.', 'факультет экономики и права',\
            'емф', 'техн.', 'фак.', 'пед.', 'на базе впо']
        for i in self.dirs:
            i = i.lower()
            for j in del_elements:
                i = i.replace(j, '')
            i = ' '.join(i.split())
            #i = re.sub("^\s+|\n|\r|\s+$", '', i)
            #i = re.sub('\s+', ' ', i)
            dirs.append(i)
        self.dirs = dirs

    def create_direction(self):
        dirs = []
        dirs1 = list(self.dirs)
        dirs1 = list(set(dirs1))
        dirs1.sort()
        for i in dirs1:
            string = i.split(' - ')
            if len(i.split(' - ')) == 2:
                dirs.append(string[1])
            elif len(i.split(' - ')) > 2:
                for i in range(1, len(string) - 1):
                    tmp = string[i].split(' [')
                    tmp = tmp[0]
                    dirs.append(tmp)
        self.dirs1 = list(set(dirs))
        self.dirs1.sort()

        '''
        for enum, element in enumerate(self.dirs):
            for i in range(enum + 1, len(self.dirs)):
                self.dirs[i] = self.dirs[i].replace(element, '')
        for i in self.dirs:
            if i != "":
                dirs.append(i)
        
        for i in dirs:
            print(i)'''
    def count(self):
        counted = {}
        for i in self.dirs1:
            counted[i] = 0
        for i in self.dirs1:
            for j in self.dirs:
                if i in j:
                    counted[i] += 1
        
        for i in counted.keys():
            print(i, counted[i])
        print('Total:',sum(counted.values()))




if __name__ == "__main__":
    counter()
