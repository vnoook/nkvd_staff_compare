# надо сравнить два файла xls файла по колонкам фио
# за главную информацию взять информацию из файла 1С

import openpyxl

# файлы данных
file_staff_1c = 'staff_1c.xlsx'
file_staff_aduser = 'staff_aduser.xlsx'

# подключаюсь к файлам, собираю инфу с нужных колонок в разные списки, сравниваю фио и в файле 1с отмечаю цветом
wb_1c = openpyxl.load_workbook(file_staff_1c)
wb_1c_s = wb_1c.active
# print(wb_1c, wb_1c_s)

# ищу минимальные и максимальные места, где есть данные
wb_1c_s_row_begin = wb_1c_s.min_row
wb_1c_s_row_end = wb_1c_s.max_row
wb_1c_s_col_begin = wb_1c_s.min_column
wb_1c_s_col_end = wb_1c_s.max_column

# читаю данные из файла
for number in wb_1c_s.iter_cols(min_col=wb_1c_s_col_begin, max_col=wb_1c_s_col_end,
                                min_row=wb_1c_s_row_begin, max_row=wb_1c_s_row_end,
                                values_only=True):
    print(number)

wb_aduser = openpyxl.load_workbook(file_staff_aduser)
wb_aduser_s = wb_aduser.active
print(wb_aduser, wb_aduser_s)
