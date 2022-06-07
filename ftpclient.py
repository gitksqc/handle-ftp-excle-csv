#!/bin/env python3
# -*- coding: utf-8 -*-

import csv
from collections import namedtuple
from decimal import Decimal
from sqlite3 import Row
import xlrd

excel = xlrd.open_workbook('sample.xls')
sheet = excel.sheets()[0]
print(sheet.nrows)
print(sheet.ncols)
# for i in range(sheet.nrows):
    # for j in range(sheet.ncols):
        # print(sheet[6][0])
print(len(sheet[6][0]))

d = '2.028'
dd = float(d)
print(round(dd, 2))
print(2 == 2.00)
# round()
# for i in range(2, 11):
capacity_avg_cycle = set([2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
print(len(capacity_avg_cycle))

with open('a.csv') as f:
    fcsv = csv.reader(f)
    # header = next(fcsv)
    Row = namedtuple('ow', next(fcsv))

    i = 0
    for row in fcsv:
        info = Row(*row)
        i += 1
        # if i < 5:
            # print(row)
            # print(info)
exit(0)

from ftplib import FTP

ftp = FTP('192.168.56.1', 'user', '123123')
ftp.encoding = 'gbk'
ftp.cwd('b/bb')
file_dirs = ftp.mlsd()
files = dict()
for s in file_dirs:
    print(s)
    files[s[1].get('modify')] = s[0]
    # print (s[0], s[1].get('type'))
print(files)
print(files[[v for v in sorted(files.keys(), reverse=True)][0]])
# ftp.cwd('/')
# file_dirs = ftp.mlsd()
# for s in file_dirs:
#     print (s[0], s[1].get('type'))

# download
# ftp.retrbinary('RETR note.txt', open('localnote.txt', 'wb').write)

# ftp.cwd('test')

# ftp.retrlines('LIST')


