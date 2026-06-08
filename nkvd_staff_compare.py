# надо сравнить два файла xls по колонкам фио и подразделение
# за главную информацию взять информацию из 1С

import openpyxl

file_staff_1c = 'staff_1c.xlsx'
file_staff_aduser = 'staff_aduser.xlsx'

# подключаюсь к файлам, собираю инфу с нужных колонок в разные списки, сравниваю фио и в файле 1с отмечаю цветом
wb_1c = openpyxl.Workbook()
wb_1c_s = wb_1c.active
print(wb_1c, wb_1c_s)
wb_1c_s1 = wb_1c_s.min_column
wb_1c_s2 = wb_1c_s.min_row
wb_1c_s3 = wb_1c_s.max_column
wb_1c_s4 = wb_1c_s.max_row
print(wb_1c_s1, wb_1c_s2, wb_1c_s3, wb_1c_s4)

wb_aduser = openpyxl.Workbook()
wb_aduser_s = wb_aduser.active
print(wb_aduser, wb_aduser_s)
