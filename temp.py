# -*- coding: utf-8 -*-
"""
Éditeur de Spyder

Ceci est un script temporaire.
"""
from Procedures import geocode2
import pandas as pd
from Procedures import correct_zip
da = pd.read_excel('C:\HomeDepot_Excel_Files\Standard_File.xlsx', sheetname='DA_List')
da['Zip_Code'] = da.apply(lambda row: correct_zip(str(row['Zip_Code'])), axis =1)
da['State2'] = da.apply(lambda row: geocode2(str(row['Zip_Code']))[0],axis = 1)
test = da.groupby(['Zip_Code','Carrier','State'])

