#!/usr/bin/env python
# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate
import xlrd
from tkinter.filedialog import askopenfilename


def template_paste(template, fileout, reg_num, names):  # (шаблон, результат, рег.номер, имена)
    doc = DocxTemplate(template)
    context = {'regnum': reg_num, 'fio': names}  # В словаре key это тег, значение - строка для вставки
    doc.render(context)  # Вставка строк из словаря в документ по тегам вида {{tag}}
    doc.save(fileout)


xls = askopenfilename()
rb = xlrd.open_workbook(xls)
sheetR = rb.sheet_by_index(0)

regNumber = set()
for row in range(sheetR.nrows):
    if sheetR.cell_type(row, 0) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
        k = str(sheetR.cell_value(row, 0))
        regNumber.add(k)

dictRegNum = {}
for reg in regNumber:
    dictRegNum[reg] = ''

for reg in regNumber:
    cnt = 1
    for row in range(sheetR.nrows):
        if reg == str(sheetR.cell_value(row, 0)):  # str, потому что в Excel это знчение Float
            pF = sheetR.cell_value(row, 1).title()
            pI = sheetR.cell_value(row, 2).title()
            pO = sheetR.cell_value(row, 3).title()
            cF = sheetR.cell_value(row, 4).title()
            cI = sheetR.cell_value(row, 5).title()
            cO = sheetR.cell_value(row, 6).title()
            pcFIO = str(cnt) + '. ' + pF + ' ' + pI + ' ' + pO + ', ' + cF + ' ' + cI + ' ' + cO
            dictRegNum[reg] = dictRegNum.get(reg) + '\n' + pcFIO
            cnt += 1
"""  (NBSB) это неразрыный пробел, чтобы Word не делал перенос по словам, вставляется по Alt+255"""
filetempl = askopenfilename()
for key in dictRegNum:
    out = key[:-2] + '.docx'
    print(key[:-2], dictRegNum.get(key))
    template_paste(filetempl, out, key[:-2], dictRegNum.get(key))
