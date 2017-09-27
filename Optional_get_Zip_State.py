# -*- coding: utf-8 -*-
"""
Created on Sat Aug 12 13:31:18 2017

@author: Bastide
"""

import openpyxl as xl
from Procedures import cell, instance
from uszipcode import ZipcodeSearchEngine
search = ZipcodeSearchEngine()
wb = xl.load_workbook('Excel Files\Standard_File.xlsx')
ws = wb['Zip_Allocation_and_Pricing']

l = instance(ws)

for r in range(l):
    if cell(ws,r+2,2) is None:
        ws.cell(row= r+2, column = 2).value = search.by_zipcode(cell(ws,r+2,1))['State']
wb.save('Excel Files\Standard_File.xlsx')