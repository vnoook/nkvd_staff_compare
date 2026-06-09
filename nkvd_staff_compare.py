# надо сравнить два файла xls файла по колонкам фио
# за главную информацию взять информацию из файла 1С

import openpyxl

def str_contains_digit(sequence):
    return any(character.isdigit() for character in sequence)

# файлы данных
file_staff_1c = 'staff_1c.xlsx'
file_staff_aduser = 'staff_aduser.xlsx'

# подключаюсь к файлу
wb_1c = openpyxl.load_workbook(file_staff_1c)
wb_1c_s = wb_1c.active

# ищу минимальные и максимальные места, где есть данные
wb_1c_s_row_begin = 1  # wb_1c_s.min_row
wb_1c_s_row_end = wb_1c_s.max_row
wb_1c_s_col_begin = 1  # wb_1c_s.min_column
wb_1c_s_col_end = wb_1c_s.max_column

# читаю данные из файла
for fio_list in wb_1c_s.iter_cols(min_col=wb_1c_s_col_begin, max_col=wb_1c_s_col_end,
                                  min_row=wb_1c_s_row_begin, max_row=wb_1c_s_row_end,
                                  values_only=True):
    pass
wb_1c.close()

list_fio_1c = []
for fio in fio_list:
    if not str_contains_digit(str(fio)):
        list_fio_1c.append(''.join(str(fio).strip().lower().split()))

len_list_fio_1c = len(list_fio_1c)
set_fio_1c = set(list_fio_1c)
len_set_fio_1c = len(set_fio_1c)

# подключаюсь к файлу
wb_aduser = openpyxl.load_workbook(file_staff_aduser)
wb_aduser_s = wb_aduser.active

# ищу минимальные и максимальные места, где есть данные
wb_aduser_s_row_begin = 1  # wb_aduser_s.min_row
wb_aduser_s_row_end = wb_aduser_s.max_row
wb_aduser_s_col_begin = 1  # wb_aduser_s.min_column
wb_aduser_s_col_end = wb_aduser_s.max_column

# читаю данные из файла
for fio_list in wb_aduser_s.iter_cols(min_col=wb_aduser_s_col_begin, max_col=wb_aduser_s_col_end,
                                  min_row=wb_aduser_s_row_begin, max_row=wb_aduser_s_row_end,
                                  values_only=True):
    pass
wb_aduser.close()

list_fio_aduser = []
for fio in fio_list:
    if not str_contains_digit(str(fio)):
        list_fio_aduser.append(''.join(str(fio).strip().lower().split()))

len_list_fio_aduser = len(list_fio_aduser)
set_fio_aduser = set(list_fio_aduser)
len_set_fio_aduser = len(set_fio_aduser)

print(set_fio_1c)
print(set_fio_aduser)

a = {"1","2","3","4","5","6"}
b = {"22","2","3","9","11","6"}
