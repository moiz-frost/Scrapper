import random
import xlrd
from xlsxwriter import Workbook

list = 'Wajeeh', (11, (11, 'Wajeeh'), 78), 2

print(list[1][1][1])

new_list = [random.randint(0, 100) for value in range(0, 30)]

print(len(new_list))

# WRITE

#Create excel file
workbook = Workbook('boys.xlsx')

#Create worksheet
worksheet_one = workbook.add_worksheet()
worksheet_two = workbook.add_worksheet()

#Write in worksheet one
worksheet_one.write(0, 0, 'Wajeeh')
worksheet_one.write(0, 1, 'Ahmed')

#Write in worksheet one
worksheet_two.write(0, 0, 'Slim')
worksheet_two.write(0, 1, 'Ansari')

workbook.close()

# READ

#Open file
workbook = xlrd.open_workbook('Names.xlsx')

#Select sheet
worksheet_one = workbook.sheet_by_index(2)

#Iterate
for row in range(worksheet_one.nrows):
    fname, lname, gender = worksheet_one.row_values(row)
    print(fname, '  ', lname, '  ', gender)

