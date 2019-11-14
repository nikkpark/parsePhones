from urllib.request import urlopen
from bs4 import BeautifulSoup
import csv
import openpyxl


starting_page = 'https://place-for-your-site.com'

def getPage(raw):
    return urlopen(raw)

def parsePage(html):
    return BeautifulSoup(html, 'lxml')

def collectHeaders(bsObj):
    h3_raw_list = bsObj.findAll('h3')
    h3_list = []

    for header in h3_raw_list:
        h3_list.append([header.text.upper()])
    return h3_list

def collectData(bsObj):
    tbody_raw_list = bsObj.find('section').findAll('tbody')
    data_list = []
    td_raw_list = []

    for i in range(len(tbody_raw_list)):
        td_raw_list.append(tbody_raw_list[i].findAll('td'))

    for lst in range(len(td_raw_list)):
        for i in range(len(td_raw_list[lst])):
            tmp = []
            tmp.append(td_raw_list[lst][i].text.strip().replace('\n', ' | '))
            data_list.append(tmp)
            tmp = []

    return data_list

def organizeData(data):
    counter = 0
    gruppen_data = []
    for line_num in range(len(data)):
        if line_num == counter:
            gruppen_data.append([data[counter], data[counter+1], data[counter+2]])
            counter = counter + 3
        else:
            continue


    return gruppen_data

def writeCsv(headers, data):
    with open('phonebook.csv', 'w',  newline='') as file:
        writer = csv.writer(file)
        counter = 0
        for i in range(len(data)):
            if data[i][0][0].startswith('ФИО'):
                writer.writerow('')
                writer.writerow(headers[counter])
                counter +=1
                continue
            else:
                writer.writerow((data[i][0][0], data[i][1][0], data[i][2][0]))

def writeXmlx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.column_dimensions['A'].width = 84
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15

    with open('phonebook.csv') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            ws.append(row)
    wb.save('phonebook.xlsx')

def debugTheBug(*args):
    for i in args:
        print(i)

def run():
    bs = parsePage(getPage(starting_page))
    data = organizeData(collectData(bs))
    headers = collectHeaders(bs)
    writeCsv(headers, data)
    writeXmlx()
   #debugTheBug(headers, data)


if __name__ == '__main__':
    run()
