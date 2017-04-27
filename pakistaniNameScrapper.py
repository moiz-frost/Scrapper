from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from xlsxwriter import Workbook
import requests
import re
import xlrd


#Output Excel File
output_excel_file = 'Muslim Names.xlsx'



boys_names = []
girls_names = []

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

# Afghan Names Loop
for pageno in range(1, 45):
    print("{0} {1}".format(pageno, "Afghan"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/afghan-baby-names.15/?order=name_date&direction=desc&category=15&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Algerian Names Loop
for pageno in range(1, 52):
    print("{0} {1}".format(pageno, "Algerian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/algerian-baby-names.13/?order=name_date&direction=desc&category=13&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Arabic Names Loop
for pageno in range(1, 302):
    print("{0} {1}".format(pageno, "Arabic"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/arabic-baby-names.4/?order=name_date&direction=desc&category=4&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Cool Names Loop
for pageno in range(1, 123):
    print("{0} {1}".format(pageno, "Cool"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/cool-baby-names.14/?order=name_date&direction=desc&category=14&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Egyptian Names Loop
for pageno in range(1, 117):
    print("{0} {1}".format(pageno, "Egyptian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/egyptian-baby-names.12/?order=name_date&direction=desc&category=12&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Farsi Names Loop
for pageno in range(1, 77):
    print("{0} {1}".format(pageno, "Farsi"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/farsi-baby-names.8/?order=name_date&direction=desc&category=8&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Indian Names Loop
for pageno in range(1, 24):
    print("{0} {1}".format(pageno, "Indian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/indian-baby-names.16/?order=name_date&direction=desc&category=16&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Iranian Names Loop
for pageno in range(1, 122):
    print("{0} {1}".format(pageno, "Iranian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/iranian-baby-names.7/?order=name_date&direction=desc&category=7&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Iraqi Names Loop
for pageno in range(1, 78):
    print("{0} {1}".format(pageno, "Iraqi"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/iraqi-baby-names.11/?order=name_date&direction=desc&category=11&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Lebanese Names Loop
for pageno in range(1, 31):
    print("{0} {1}".format(pageno, "Lebanese"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/lebanese-baby-names.10/?order=name_date&direction=desc&category=10&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Malaysian Names Loop
for pageno in range(1, 66):
    print("{0} {1}".format(pageno, "Malaysian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/malaysian-baby-names.9/?order=name_date&direction=desc&category=9&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Muslim Names Loop
for pageno in range(1, 386):
    print("{0} {1}".format(pageno, "Muslim"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/muslim-baby-names.3/?order=name_date&direction=desc&category=3&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Pakistani Names Loop
for pageno in range(1, 313):
    print("{0} {1}".format(pageno, "Pakistani"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/pakistani-baby-names.1/?order=name_date&direction=desc&category=1&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Pashtun Names Loop
for pageno in range(1, 10):
    print("{0} {1}".format(pageno, "Pashtun"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/pashtun-baby-names.19/?order=name_date&direction=desc&category=19&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Persian Names Loop
for pageno in range(1, 110):
    print("{0} {1}".format(pageno, "Persian"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/persian-baby-names.5/?order=name_date&direction=desc&category=5&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Punjabi Names Loop
for pageno in range(1, 7):
    print("{0} {1}".format(pageno, "Punjabi"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/punjabi-baby-names.6/?order=name_date&direction=desc&category=6&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Quranic Names Loop
for pageno in range(1, 94):
    print("{0} {1}".format(pageno, "Quranic"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/quranic-baby-names.21/?order=name_date&direction=desc&category=21&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

# Sahabah Names Loop
for pageno in range(1, 4):
    print("{0} {1}".format(pageno, "Sahabah"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/sahabah-names.22/?order=name_date&direction=desc&category=22&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

#Sindhi Names Loop
for pageno in range(1, 4):
    print("{0} {1}".format(pageno, "Sindhi"))
    # URL
    url = 'https://www.pakistan.web.pk/names/category/sindhi-baby-names.17/?order=name_date&direction=desc&category=17&name_state=visible&page={0}'.format(pageno)

    # Get the webpage object with 9 sec timeout
    responseObject = requests.get(url, headers=headerForBrowserRequests, timeout=9)

    # Store the HTML source code in output file
    with open(filename, 'w+') as output:
        output.write(responseObject.content) if type(responseObject.content) == bytes else output.write(
            responseObject.content)

    # Create soup object
    soup = BeautifulSoup(read_file(), 'lxml')

    # Append names to the list
    girls_divs_names = soup.find_all('img', attrs={'title': 'Baby Girl Names', 'src': 'styles/default/babynames/gender/f_s.png'})
    boys_divs_names = soup.find_all('img', attrs={'title': 'Baby Boy Names', 'src': 'styles/default/babynames/gender/m_s.png'})

    for girl in girls_divs_names:
        parent = girl.parent.parent.parent
        girls_names.append(parent.h3.a.contents[0])

    for boy in boys_divs_names:
        parent = boy.parent.parent.parent
        boys_names.append(parent.h3.a.contents[0])

#19386


print(boys_names)
print(girls_names)

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
