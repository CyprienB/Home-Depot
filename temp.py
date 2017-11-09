# -*- coding: utf-8 -*-
"""
Éditeur de Spyder

Ceci est un script temporaire.
"""
from Procedures import geocode2, geocode3, most_common
import pandas as pd
from Procedures import correct_zip
from geopy.geocoders import Nominatim, GoogleV3, GeocoderDotUS

da = pd.read_excel('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx', sheetname='Zip', converters={'Latitude': float,'Longitude': float})

da['ZipCode'] = da['ZipCode'].apply(lambda row: correct_zip(str(row)))
da['State'] = da.apply(lambda row: geocode3(str(row['ZipCode']),row['Latitude'],row['Longitude']),axis = 1)

#postal = '00601'
#lat = 18.180555
#long = -66.749961
##info = GoogleV3().geocode(str(postal)+", United States of America")    
##a = info.raw
#
##info1 = geocode3(postal,18.180555,-66.749961)
#info = GoogleV3().reverse('%d, %d' %(lat, long))   
#print(info)
#state = []
#
#for i in info:
#    for string in us_state_abbrev.keys():
#        if i[0].find(string) != -1:
#            state.append(us_state_abbrev[string])
