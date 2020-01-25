import db_ege
import db_users
import db_training_dir
import xlrd

if __name__ == "__main__":
    workbook = xlrd.open_workbook('123.xls', encoding_override='utf-8')
    worksheet = workbook.sheet_by_index(0)
    db_users.parsing_users(worksheet)
    db_ege.parsing_exams(worksheet)    
    db_training_dir.parsing_training_dir(worksheet)