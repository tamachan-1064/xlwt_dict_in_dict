# -*- coding: utf-8 -*-
import xlwt
import os

alist = ['a1', 'a2','a3']
klist = ['k1', 'k2','k3']
d = {'a1':{'k1': 1, 'k2':2, 'k3': 3}, 'a2':{'k1': 4, 'k2':5, 'k3': 6}, 'a3':{'k1': 7, 'k2':8, 'k3': 9}}

wb = xlwt.Workbook()

sheet = wb.add_sheet('sheet1', cell_overwrite_ok=True)

for i, key1 in enumerate(alist):
    for j, key2 in enumerate(klist):
        sheet.write(i+1, 0, str(key1))
        sheet.write(0, j+1, str(key2))
        sheet.write(i+1, j+1, d[key1][key2])

cwd = os.getcwd()

wb.save(cwd +'/xlwt_dict_in_dict.xls')