# -*- coding: utf-8 -*-
"""
Created on Fri Aug  4 14:42:10 2017

@author: Bastide
"""

def geocode2(postal, recursion=0):
#    Returns the list[City,Postal Code,(lat,long)
    from geopy.exc import GeocoderTimedOut
    from geopy.geocoders import GoogleV3
    import time

#Look Online
    try:
        info = GoogleV3().geocode(str(postal)+", United States of America")    
        #       Avoid time out error
    except GeocoderTimedOut as e:
        if recursion > 10:      # max recursions
            raise e
        time.sleep(0.1) # wait a bit
        # try again
        return geocode2(postal,recursion=recursion + 1)
    else: 
        # If no result found use previous info
        if info is None:
            return None

        else :
            city = info.address
            lat=info.latitude
            long=info.longitude
            time.sleep(0.1)
#            state= info.state
            return [city,
                postal,
                (lat,long)]
#                ,
#                state]


# Facilitate the openpyxl formating
def cell(Sheet, rownb, columnnb):  
    return Sheet.cell(row=rownb, column=columnnb).value

# Return the number of instance in the Sheet, we can adjust the starting line
def instance(Sheet, starting_row=2, column=1):
    r=0
    while cell(Sheet, r+starting_row, column) is not None:
        r+=1
    return r


# Return a dictionnary of pricing (State : Carrier : ( Flat , Break, Extra))
def get_lm_pricing(Sheet):
    Pricing={}
# Get State in DIct
    nb_state = instance(Sheet,starting_row=3)
    for r in range(nb_state):
        Pricing[cell(Sheet,3+r,2)]= {}
# Get carriers
    c=0
    Carriers = []
    while cell(Sheet,1,4+3*c) is not None:
        Carriers.append(cell(Sheet,1 ,4+3*c))
        c+=1
# Create Dictionnaries
    for r in range(len(Pricing)):
        for c in range(len(Carriers)):
#            Append only if there is pricing info
            if cell(Sheet,r+3,3*c+3) is not None:
                Pricing[cell(Sheet, r+3,2)][Carriers[c]]={'Flat':cell(Sheet,r+3,3+3*c),'Break':cell(Sheet,r+3,4+3*c),'Extra':cell(Sheet,r+3,5+3*c)}
    return Pricing

#    Return a list of all"neighboring" states to the one in the argument 
def neig_states(state_code, Sheet):
    a=instance(Sheet)
    List=[state_code]
    for r in range(a):
        if cell(Sheet,r+2,1) == state_code:
            List.append(cell(Sheet,r+2, 2))
            
        if cell(Sheet, r+2, 2) == state_code and cell(Sheet,r+2,3) == "1st":
            List.append(cell(Sheet, r+2, 1))
    return List
    
# Return a list of STates defining The region, based on degree of neighboor 
def get_second_neig(state_code, Sheet):
    List = neig_states(state_code, Sheet) 
    List2 = List
    for state in List:
        A = neig_states(state,Sheet)
#       Remove Duplicates
        List2 = list(set().union(List2,A))
    return List2          

def compute_distance2(zip1,zip2,Dict_lat_long):
    b = 0
    from geopy.distance import vincenty
    from Procedures import geocode2    
#   a serve to know if zipcode not in database appears and will
    try:
        latlong1 = Dict_lat_long[zip1]
    except KeyError:
        Dict_lat_long[zip1] = geocode2(zip1,10)[2]
        latlong1 = Dict_lat_long[zip1]
        b=1
        
    try:
        latlong2 = Dict_lat_long[zip2]
    except KeyError:
        Dict_lat_long[zip2] = geocode2(zip2,10)[2]
        latlong2 = Dict_lat_long[zip2]    
        b=1
        
    distance = vincenty(latlong1,latlong2).miles
    
    return distance, Dict_lat_long ,b

          
# This function will return a 5 digit postal code by adding 0 in front if the input is less than 5
def correct_zip(str_Zip):
    if len(str_Zip) == 4:
        zipcode = "0"+str_Zip
    elif len(str_Zip) == 3:
        zipcode = "00"+str_Zip
    elif len(str_Zip)>5:
        a = len(str_Zip)-5
        zipcode = str_Zip[a:]
    else :
        zipcode = str_Zip
    return zipcode
    
# This function will return the dictionary of every State Destination with weighted origin 
def averageOrig (ltl_price):
    import pandas as pd
    
    destWeight = pd.DataFrame({'total' : ltl_price.groupby(['dest_state'])['sys_invc_id'].count()}).reset_index()
    origWeight = pd.DataFrame({'subtotal' : ltl_price.groupby(['dest_state','orig_state'])['sys_invc_id'].count()}).reset_index()
    origWeight = origWeight.merge(destWeight[['dest_state','total']], on=['dest_state'])
    origWeight['percentage'] = round(origWeight['subtotal']/origWeight['total'],4)

    percDestin = {}
    for row in origWeight.iterrows():
        percDestin.setdefault(row[1]['dest_state'],{}).setdefault(row[1]['orig_state'],row[1]['percentage'])

    return percDestin 