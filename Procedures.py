# -*- coding: utf-8 -*-
"""
Created on Fri Aug  4 14:42:10 2017

@author: Bastide
"""

def geocode2(postal, recursion=0):
#    Returns the list[City,Postal Code,(lat,long)

    state_db = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DC", "DE", "FL", "GA", 
          "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", 
          "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", 
          "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", 
          "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"]

    from geopy.exc import GeocoderTimedOut
    from geopy.geocoders import GoogleV3, Nominatim
    import time

#Look Online
    try:
        info = GoogleV3(timeout = 10).geocode(str(postal)+", United States of America")    
        #       Avoid time out error
    except GeocoderTimedOut as e:
        if recursion > 10:      # max recursions
             return ['unknown','unknown',('unknown','unknown'),'unknown']
        time.sleep(0.1) # wait a bit
        # try again
        return geocode2(postal,recursion=recursion + 1)
    except :
        info = None

    # If no result found use previous info
    if info is None:
        return ['unknown','unknown',('unknown','unknown'),'unknown']

    else :
        a = info.raw['address_components']
        
        for row in a:
                try: 
                    b = row['short_name']
                    
                except:
                    b = 0
                if b in state_db:
                    state = b
                    
                    break
                else:
                    state = 'unknown'    
        city = 'city is not looked for'
        lat=info.latitude
        long=info.longitude
        time.sleep(0.5)
#            state= info.state
        return [city,
            postal,
            (lat,long),
            state]


def geocode3(postal,lat,long, recursion=0):
#    Returns the list[City,Postal Code,(lat,long)
    state_db = us_state_abbrev = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY',
    'Puerto Rico': 'PR'
}
    from geopy.exc import GeocoderTimedOut
    from geopy.geocoders import GoogleV3, Nominatim
    import time
    import uszipcode as usz
    
    search = usz.ZipcodeSearchEngine()
    info = search.by_zipcode(correct_zip(str(postal)))
    
    if info["State"] is not None:
        state = info["State"]
    else:
        state = 'None'
#    #Look Online
#        try:
#            info = GoogleV3().reverse('%d, %d' %(lat, long))    
#            
#            #       Avoid time out error
#        except GeocoderTimedOut as e:
#            if recursion > 10:      # max recursions
#                raise e
#            time.sleep(0.1) # wait a bit
#            # try again
#            return geocode3(lat,long,recursion=recursion + 1)
#        else: 
#            # If no result found use previous info
#            if info is None:
#                return None
#    
#            else :
#                state = []
#                for i in info:
#                    for string in us_state_abbrev.keys():
#                        if i[0].find(string) != -1:
#                            state.append(us_state_abbrev[string])
#                state = most_common(state) 
    print(state)
    return state

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
    a= len(Sheet)
    List=[state_code]
    for r in range(a):
        if Sheet['StateCode'][r] == state_code:
            List.append(Sheet['NeighborStateCode'][r])
            
        if Sheet['NeighborStateCode'][r] == state_code and Sheet['Neighboring Degree'][r] == "1st":
            List.append(Sheet['StateCode'][r])
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

def compute_distance2(zip1,zip2,Dict_lat_long, show =0,Minkowski_coef = 1.54):
    b = 0
    from geopy.distance import vincenty
    from Procedures import geocode2    
#   a serve to know if zipcode not in database appears and will
    try:
        latlong1 = Dict_lat_long[zip1][0]
    except KeyError:
        info = geocode2(zip1)
        latlong1 = info[2]
        state = info[3]
        Dict_lat_long[zip1] = [latlong1, state]
        b=1
        
    try:
        latlong2 = Dict_lat_long[zip2][0]
    except KeyError:
        info = geocode2(zip2)
        latlong2 = info[2]
        state = info[3]
        Dict_lat_long[zip2] = [latlong2, state]
        b = 1
    if show ==1:
        print(latlong1,latlong2)
    if latlong1 == ('unknown','unknown') or latlong2 == ('unknown','unknown'):
        distance = 1000
    else:
        distance = vincenty(latlong1,latlong2).miles * Minkowski_coef
    return distance, Dict_lat_long ,b

          
# This function will return a 5 digit postal code by adding 0 in front if the input is less than 5
def correct_zip(str_Zip):
    str_Zip = str(str_Zip)
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


def most_common(L):
    
    import itertools
    import operator
# get an iterable of (item, iterable) pairs
    SL = sorted((x, i) for i, x in enumerate(L))
# print 'SL:', SL
    groups = itertools.groupby(SL, key=operator.itemgetter(0))
# auxiliary function to get "quality" for an item
    def _auxfun(g):
        item, iterable = g
        count = 0
        min_index = len(L)
        for _, where in iterable:
            count += 1
            min_index = min(min_index, where)
# print 'item %r, count %r, minind %r' % (item, count, min_index)
        return count, -min_index
# pick the highest-count/earliest item
    return max(groups, key=_auxfun)[0]