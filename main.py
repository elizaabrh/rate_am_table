import os
import io
import time
import requests
import schedule

from bs4 import BeautifulSoup
from openpyxl import Workbook
from typing import List

workbook = Workbook()
sheet = workbook.active


def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start + len(needle))
        n -= 1
    return start


def get_table():
    r = requests.get("https://www.rate.am")
    page = r.text
    soup = BeautifulSoup(page, 'html.parser')
    tables = soup.findChildren('table')

    my_table = tables[3]
    rows = my_table.findChildren(['th', 'tr'])

    i = 1
    j = 68
    for option in my_table.find_all('option'):
        option = str(option)
        if "selected" in option:
            # for i in range(1, 20):
            sheet[chr(j) + str(i)] = sheet[chr(j + 1) + str(i)] = option[option.index(">") + 1:option.index(">") + 6]
            j += 2

    i = 2
    j = 65
    for a in my_table.find_all('a'):
        a = str(a)
        if "Դասակարգել" in a:
            sheet[chr(j) + str(i)] = a[a.index('>') + 1:find_nth(a, "<", 2)]
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

    # for a in my_table.find_all('a'):
    #     print(a)

    print(rows)


# for line in var:
#     for j in range(65, 78):
#         for i in range(1, 20):
#             sheet[chr(j) + str(i)] = var


schedule.every(5).minutes.do(get_table)

get_table()

workbook.save(filename="data.xlsx")
os.system('start excel.exe data.xlsx')

# while True:
#     schedule.run_pending()
#     time.sleep(5)
