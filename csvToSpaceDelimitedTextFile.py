import csv

names = []
clean_names = []

# Input CSV File
csvfile = open('Comma.csv', 'r')
# Output Text File
txt = open('out.txt', 'a')


csv = csv.reader(csvfile, delimiter=',')

i = 0

for row in csv:
    names.append(row)

for value in names:
    clean_names.append(value[0].strip('\n'))


for name in clean_names:
    print(name)
    txt.write('{0} '.format(name))

txt.close()
csvfile.close()
