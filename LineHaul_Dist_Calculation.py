# -*- coding: utf-8 -*-
"""
Created on Tue Oct 24 15:41:35 2017

@author: Steven Gao
"""

#from Procedures import compute_distance2,cell, instance
import pandas as pd
#tqdm is used to create progress bar, however, if computation is too fast it might cause bugs
from tqdm import tqdm
#This Module is used to open excel files
import openpyxl as xl
from Procedures import cell, instance, compute_distance2, correct_zip


# Import origin and destinaiton zip codes from ltl price file
ltl_price = pd.read_excel('C:\HomeDepot_Excel_Files\ltl_price.xlsx',sheetname='ltl_price')

# Import Database of Zipcode Latitude and Longitude
wdata = xl.load_workbook('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
wslatlong = wdata['Zip']

distance_list = []
orig_zip = []
dest_zip = []
orig_zip_data = ltl_price.iloc[0:len(ltl_price), 18]
dest_zip_data = ltl_price.iloc[0:len(ltl_price), 23]


#Change Data Structure from dataframe to lists
for i in range(0, len(orig_zip_data)):
    orig_zip.append(orig_zip_data.get_value(i))
    dest_zip.append(dest_zip_data.get_value(i))

# Create dictionnary for the database lat long{Zip : (lat,long)}
linelatlong = instance(wslatlong)
Zip_lat_long = {}
for r in range(linelatlong):
    zipcode = correct_zip(str(cell(wslatlong,r+2,1)))
    lat = cell(wslatlong,r+2,2)
    long = cell(wslatlong,r+2,3)
    Zip_lat_long[zipcode] = (lat,long)

    
# Compute distances for each combination
nb_distances = len(orig_zip)

print("Compute distances")
for i in tqdm(range(nb_distances)):
    zipcode1 = orig_zip[i]
    zipcode2 = dest_zip[i]
    distance, Zip_lat_long, b = compute_distance2(zipcode1, zipcode2, Zip_lat_long)
    distance_list.append(distance)

# Output to an excel file
df = pd.DataFrame({'tot_mile_cnt': distance_list})
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\LH_Dist_Recalc.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

