import xlrd
from xlsxwriter import Workbook

names_list = []
gender_list = []

read_workbook = xlrd.open_workbook('Muslim Names.xlsx')
read_workbook_sheet = read_workbook.sheet_by_index(0)

new_workbook = Workbook('CleanedCSV.xlsx')
worksheet_one = new_workbook.add_worksheet()

i = 0
# Iterate
for row in range(read_workbook_sheet.nrows):
    name, gender = read_workbook_sheet.row_values(row)
    names_list.append(name.strip('\n'))
    gender_list.append(gender.strip('\n'))
    i += 1

for i in range(0, 19787):
    worksheet_one.write(i, 0, names_list[i])
    worksheet_one.write(i, 1, gender_list[i])

new_workbook.close()

print(i)

# for value in names:
#     clean_names.append(value[0].strip('\n'))
#
# for name in clean_names:
#     print(name)
