import os
import io
import re
import time
import requests
import schedule
import pandas

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment

workbook = Workbook()
sheet = workbook.active

sheet.merge_cells('D1:E1')
sheet.merge_cells('F1:G1')
sheet.merge_cells('H1:I1')
sheet.merge_cells('J1:K1')

i = 1
j = 4
while j < 11:
    sheet.cell(i, j).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    j += 2

sheet.column_dimensions['A'].width = 25
sheet.column_dimensions['B'].width = 15
sheet.column_dimensions['C'].width = 15


def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start + len(needle))
        n -= 1
    return start


#   for j in range(65, 78):
#       for i in range(1, 20):
#           sheet[chr(j) + str(i)] = var


def get_table():
    r = requests.get("https://www.rate.am")
    page = r.text
    soup = BeautifulSoup(page, 'html.parser')
    tables = soup.findChildren('table')
    my_table = tables[3]

    i = 1
    j = 68
    for option in my_table.find_all('option'):
        option = str(option)
        if "selected" in option:
            sheet[chr(j) + str(i)] = option[option.index(">") + 1:option.index(">") + 6]
            j += 2

    i = 2
    j = 65
    for a in my_table.find_all('a'):
        a = str(a)
        if "Դասակարգել" in a:
            sheet[chr(j) + str(i)] = a[a.index('>') + 1:find_nth(a, "<", 2)]
            if "<br/>" in a:
                b = a[a.index('>') + 1:find_nth(a, "<", 2)]
                sheet[chr(j) + str(i)] = b + a[find_nth(a, ">", 2) + 1:find_nth(a, "<", 3)]
            j += 1

    i = 3
    j = 65
    for a in my_table.find_all('a'):
        a = str(a)
        if "անկ" in a or "ԱՆԿ" in a:
            if "Դասակարգել" not in a and "class" not in a:
                sheet[chr(j) + str(i)] = a[a.index('>') + 1:find_nth(a, "<", 2)]
                i += 1

    i = 3
    j = 66
    for a_href in my_table.find_all('td'):
        a_href = str(a_href)
        if "a href" in a_href and "bank" in a_href:
            if "class" not in a_href:
                sheet[chr(j) + str(i)] = a_href[find_nth(a_href, ">", 2) + 1:find_nth(a_href, "<", 3)]
                i += 1

    i = 3
    j = 67
    for date in my_table.find_all('td'):
        date = str(date)
        if "class=\"date\"" in date:
            sheet[chr(j) + str(i)] = date[find_nth(date, ">", 1) + 1:find_nth(date, "<", 2)]
            i += 1

    data = soup.find_all('td')
    numbers = [d.text for d in data if
               d.text.isdigit() or not len(d.text) or re.match(r'^[-+]?\d+(?:\.\d+)$', d.text)]

    numbers.pop(0)
    i = 191
    while i >= 153:
        numbers.pop(i)
        i -= 1

    f = 144
    while f >= 0:
        numbers.pop(f)
        f -= 9

    h = 0
    for i in range(3, 20):
        for j in range(68, 76):
            sheet[chr(j) + str(i)] = numbers[h]
            h += 1


schedule.every(5).minutes.do(get_table)

get_table()

workbook.save(filename="data.xlsx")
# os.system('start excel.exe data.xlsx')

# while True:
#     schedule.run_pending()
#     time.sleep(5)
