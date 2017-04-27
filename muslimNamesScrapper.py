from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from xlsxwriter import Workbook
import requests
import re
import xlrd

boys_names = []
girls_names = []


# Output Excel File
output_excel_file = 'Muslim Names.xlsx'




# Output for HTML code
filename = "output.txt"


def read_file():
    file = open(filename)
    data = file.read()
    file.close()
    return data


# Initialize user agent to spoof browsers
userAgentForSpoofing = UserAgent()

# Creater a header dictionary
headerForBrowserRequests = {'user-agent': userAgentForSpoofing.chrome}

# Boys loop
for pageno in range(1, 266):
    # URL
    url = 'http://www.muslimnames.info/baby-boys/islamic-boys-names-{0}/'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    all_divs_names = soup.find_all('div', attrs={'class': 'nrow_name boys'})
    for div in all_divs_names:
        boys_names.append(div.b.a.string)

    print('Boys Page No: {0}, Length: {1}'.format(pageno, len(boys_names)))

# Girls loop
for pageno in range(1, 244):
    # URL
    url = 'http://www.muslimnames.info/baby-girls/islamic-girls-names-{0}/'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    all_divs_names = soup.find_all('div', attrs={'class': 'nrow_name girls'})
    for div in all_divs_names:
        girls_names.append(div.b.a.string)

    print('Girls Page No: {0}, Length: {1}'.format(pageno, len(girls_names)))

print(len(boys_names))
print(len(girls_names))

workbook = Workbook(output_excel_file)
boys_worksheet = workbook.add_worksheet()
girls_worksheet = workbook.add_worksheet()

for i in range(0, len(boys_names)):
    boys_worksheet.write(i, 0, boys_names.pop())
    boys_worksheet.write(i, 1, 'M')

for i in range(0, len(girls_names)):
    girls_worksheet.write(i, 0, girls_names.pop())
    girls_worksheet.write(i, 1, 'F')

workbook.close()
