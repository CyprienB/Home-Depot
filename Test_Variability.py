# -*- coding: utf-8 -*-
"""
Created on Fri Sep 29 14:53:36 2017
This file combine all other Python files with some improvments to avoid useless processes
This is an optimization of the Home Depot .com Delivery network, with objective function being the total cost (line haul + last Mile) 
@author: Cyprien Bastide, Steven (Gao) Ming, Edson David Silva Moreno
"""

# This module is used to import the different modules that will be used in this file
import sys
# This module is used to open excel files
import openpyxl as xl
# This module is used to perform regression analysis
from statsmodels.formula.api import ols
# This module is used to create progress bar.
# This module is used to convert spreadsheet into a specific dataframe for analysis
import pandas as pd
# neig_states: it returns the neighboring state of the input state
# compute_distance2: it computes the distances from 2 zip code and the lat long database
# correct_zip: it adds 0 in front of postal codes that are not 5 digits long
# get_lm_pricing: it returns a dictionnary containing the info to compute the last mile cost
# averageOrig: # it returns the dictionary of every State Destination with weighted origin 
from Procedures import neig_states, compute_distance2, correct_zip, get_lm_pricing, averageOrig, geocode2

nb_iter = 50

optimization_time = 1500
oportunity_threshold = 50
oportunity_cost = 25
number_days = 30*6
weight_treshold_ltl = 200
nb_trucks = round(number_days*5/7)
weight_per_volume = 180
DA_to_DA_min_distance = 40*1.54

# Import and convert spreadsheets into panda dataframe
wb = pd.ExcelFile('C:\HomeDepot_Excel_Files\Standard_File.xlsx')
wp = pd.ExcelFile('C:\HomeDepot_Excel_Files\Standard_File_Optimized.xlsx')
wd = pd.ExcelFile('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
w_neig = wb.parse('List_of_Neighboring_States')
w_da = wb.parse('DA_List')
w_zip = wb.parse('Zip_Allocation_and_Pricing')
w_opt= wp.parse('Zip_Allocation_and_Pricing')
w_range = wb.parse("Zip_Range")
w_dfc = wb.parse("DFC list")
wslatlong = wd.parse('Zip_Lat_Long')
ltl_price = wb.parse('ltl_price', converters={'dest_zip': str,'orig_zip': str})
Volume = w_zip['Volume']
Zip = w_zip['Zip#']

    
#w_zip['Volume'] = w_zip['Volume'].sample(frac=1).reset_index(drop=True)
#w_zip['Volume'] = w_zip['Volume'] *1.5 *1.5 *1.5 *1.5 *1.5


"""
###############
###############
Regression
###############
###############
"""

# Parse useful data for regression analysis
ltl_price = ltl_price[(ltl_price['tot_shp_wt'] >= 200) & (ltl_price['tot_shp_wt'] <= 4999) & (ltl_price['aprv_amt'] <=1000)]
# Define an interaction term between distance and weight
ltl_price['tot_mile_wt'] = ltl_price['tot_mile_cnt'] * ltl_price['tot_shp_wt']

# Fit this regression model with .fit() and show results
linehaul_model = ols('aprv_amt ~ tot_mile_cnt + tot_shp_wt + + tot_mile_wt', data=ltl_price).fit()

# Output the result of the regression model
linehaul_model_summary = linehaul_model.summary()
print(linehaul_model_summary)

# Store significant terms from the regression model
variables = [linehaul_model.params.index.tolist()][0]

# Filter and rename orig_state if 'orig_state' variable is used in the regression model
for i in range(0, len(variables)):
    if variables[i].find("orig_state") != -1:
        variables[i] = variables[i][-3]+variables[i][-2]

# Store coefficents to the significant terms from the regression model
coeff = [linehaul_model.params.tolist()][0]

# Convert two lists (significant terms & coefficents) to a dictionary
coefficient = dict(zip(variables,coeff))


"""
###############################################################
###############################################################

This part create the different dictionnaries that are going to be used to solve the optimization problem
Multiple dictionnaries are going to be created to represent DAs and Zipcode because we will need different keys to call them

###############################################################
###############################################################
"""

print ("Create all Dictionnaries")


# Get length of DA, Zip, latlong, fullfillment centers, zip range
n_da= len(w_da)
n_zip = len(w_zip)
linelatlong = len(wslatlong)
nbdfc = len(w_dfc)
n_range = len(w_range)

# Create dictionnary for the database lat long{Zip : [(lat,long),state]}
Zip_lat_long = {}
for r in range(linelatlong):
    zipcode = correct_zip(str(wslatlong['ZipCode'][r]))
    lat = wslatlong['Latitude'][r]
    long = wslatlong['Longitude'][r]
    state = wslatlong['State'][r]
    Zip_lat_long[zipcode] = [(lat,long),state]
print('done')

# Dictionnary for Zip {State : [( zipcode, volume)]}
State_Zip_dict = {}
#Create Dictionnary for Zipcode (Zip, Volume ,State)
ZipCode_Dict={}
for r in range(n_zip):
    zipcode = correct_zip(str(w_zip['Zip#'][r])) 
    volume = w_zip['Volume'][r]
    try : 
        state = Zip_lat_long[zipcode][1]
    except KeyError:
        info = geocode2(zipcode)
        Zip_lat_long[zipcode] = [info[2], info[3]]
        state = Zip_lat_long[zipcode][1]
    State_Zip_dict.setdefault(state, []).append((zipcode, volume))
    ZipCode_Dict[zipcode]={'Zip':zipcode,'Volume':volume, 'State':state}
print('done')
# Dictionnary for DA  {State : [Da_ZipCode]}. This is useful because Arcs are going to be created based on neighgboring states
State_Da_dict = {} 
# Dictionnary for DA + carrier {State : [DA + carrier]}  
State_DAC_dict = {}
# Dictionnary for Da { Zip:(Zip, State, [Carrier])} will be useful to compute distances       
DA_ZipCode_Dict = {}     
# Other dictionnary for Da, { Zip + Carrier : (Zip, State, Carrier)}   
DAC_ZipCode_Dict = {}       

for r in range(n_da):
    zipcode = correct_zip(str(w_da['Zip_Code'][r]))
    carrier = w_da['Carrier'][r]
    try :
        state = Zip_lat_long[zipcode][1]
    except KeyError :
        info = geocode2(zipcode)
        Zip_lat_long[zipcode] = [info[2], info[3]]
    if Zip_lat_long[zipcode][1] == 'unknown':
        sys.exit(' Prompt to Zip_Lat_Long database Zipcode %s latitude, Longitude and State or Remove DA %s %s from DA_List' % (zipcode, zipcode, carrier))
    else : 
        state = Zip_lat_long[zipcode][1]
        

    # Remove duplicate: DA in same zipcode but different carriers
    State_Da_dict[state] = list(set().union(State_Da_dict.setdefault(state,[]),[zipcode]))
    State_DAC_dict[state] = list(set().union(State_DAC_dict.setdefault(state,[]),[zipcode +" "+ carrier]))
    DA_ZipCode_Dict.setdefault(zipcode,{'Zip':zipcode, 'State':state, 'Carrier':[carrier]}) 
    DA_ZipCode_Dict[zipcode]['Carrier'] = list(set().union(DA_ZipCode_Dict[zipcode]['Carrier'],[carrier]))
    DAC_ZipCode_Dict[zipcode+' '+ carrier] = {'Zip':zipcode, 'State':state, 'Carrier':carrier}
print('done')
# Create Dictionnary of DFC {State : {Name, Zip, State}}
DFC_Dict={}
for r in range(nbdfc):
    name = w_dfc['DFC'][r]
    zipcode = correct_zip(str(w_dfc['DFC ZIP'][r]))
    state = w_dfc['DFC State'][r]
    DFC_Dict[state]={'State':state,'Name':name,'Zipcode':zipcode}
print('done')
# Import LM Pricing again using openxl
wx = xl.load_workbook('C:\HomeDepot_Excel_Files\Standard_File.xlsx')
w_lm_pricing = wx["LM_Pricing"]

# Create LM Pricing Dictionnary
Pricing = get_lm_pricing(w_lm_pricing)

# Get arc max range Dictionnary {State : Max range}
Range = { w_range['Abreviation'][r] : w_range['Maximum distance between Zip (in state) to DA (out of state)'][r] for r in range(n_range)}

# Create dictionary with State Destination as a Key and nested dictionary with weight by origin 
percDestin = averageOrig(ltl_price)
"""
###############################################################
###############################################################

This part compute the distance between Das and DFC and the pricing based on parameters and distance

\\ Problem for new DAs in States without Das currently (Key Error)

###############################################################
############################################################### 
"""
# Relation Da_Dfc is in format {da : {state_dfc : {distance, percentage from state}, global : {slope,cost_opening}}

Da_Dfc = {}

for da in DA_ZipCode_Dict.keys():
    
    da_state = DA_ZipCode_Dict[da]["State"]
    try:
        for dfc_state in percDestin[da_state].keys():
            
            dfc_zip = DFC_Dict[dfc_state]['Zipcode']
            percentage = percDestin[da_state][dfc_state]
            
            distance, Zip_lat_long, _ = compute_distance2(da,dfc_zip,Zip_lat_long)
            
            slope = coefficient["tot_mile_wt"]+coefficient["tot_mile_wt"]*distance
            
            intercept = coefficient['Intercept']+coefficient['tot_mile_cnt']*distance
            
            cost_opening = intercept + weight_treshold_ltl * slope
            
            Da_Dfc.setdefault(da,{}).setdefault(dfc_state, {"distance":distance, "percentage" : percentage,"slope": slope,"cost_opening": cost_opening})

        slope = sum(Da_Dfc[da][dfc_state]["slope"]*Da_Dfc[da][dfc_state]["percentage"] for dfc_state in Da_Dfc[da].keys())
        
        cost_opening = sum(Da_Dfc[da][dfc_state]["cost_opening"]*Da_Dfc[da][dfc_state]["percentage"] for dfc_state in Da_Dfc[da].keys())
        
        Da_Dfc[da]['Global']={'slope' : slope, 'cost_opening' : cost_opening}
        
    except KeyError:
#        Just assume we take the previous cost
        Da_Dfc.setdefault(da,{}).setdefault('Global',{'slope' : 0.1, 'cost_opening' : 70, 'Warning':"This Da doesn't have real slope or cost of opening"}  )  

"""
##############################
Cost Analysis
##############################
##############################
"""
Current_Cost = []
Optimized_Cost = []
for iteration in range(nb_iter):
    print(iteration)
    #Shuffle Volumes for both Current and Optimized Network
    Volume_Shuffle = Volume.sample(frac=1).reset_index(drop=True)
    w_zip['Volume'] = Volume_Shuffle
    Volume_Dict = {}
#    Create Dictionnary
    for r in range(len(Volume_Shuffle)):
        zipcode = correct_zip(str(Zip[r]))
        Volume_Dict[zipcode] = Volume_Shuffle[r]
#    Assign Volume to Optimized Network
    w_opt['Volume'] = w_opt['Zip#'].apply(lambda row: Volume_Dict[correct_zip(str(row))])
    
    # Computation of LM Cost of Current Network
    Current_LM = []
    Optimized_LM = []
#    Current
    for r in range(len(w_zip)):
        zipcode = w_zip['Zip#'][r]
        if zipcode in w_opt['Zip#'].values:  # Maybe incorrect, the objective is to use only Zipcode that are assign after optimization
            zipcode = correct_zip(str(zipcode))
            da_zipcode = correct_zip(str(w_zip['DA ZIP'][r]))
            carrier = w_zip['Carrier'][r]
            state = Zip_lat_long[da_zipcode][1]
            distance, Zip_lat_long, _ = compute_distance2(zipcode,da_zipcode,Zip_lat_long)
            if carrier in Pricing[state].keys():
                flat = Pricing[state][carrier]['Flat']
                breakpoint = Pricing[state][carrier]['Break']
                extra = Pricing[state][carrier]['Extra']
                if distance <  breakpoint:
                    lmcost=flat
                else :
                    lmcost=flat+ (distance - breakpoint) * extra
               
            else : # If carrier is not in our pricing dictionnary we compute averageof all carrier present
                costs = []
                for carrier in Pricing[state].keys():
                    flat = Pricing[state][carrier]['Flat']
                    breakpoint = Pricing[state][carrier]['Break']
                    extra = Pricing[state][carrier]['Extra']
                    if distance <  breakpoint:
                        lmcost=flat
                    else :
                        lmcost=flat+ (distance - breakpoint) * extra
                    costs.append(lmcost)
                lmcost = sum(costs)/len(costs)
            Current_LM.append(lmcost)
        else:
            Current_LM.append(0)
            
    # Optimized
    for r in range(len(w_opt)):
        zipcode = w_opt['Zip#'][r]
        if zipcode in w_opt['Zip#'].values:  # Maybe incorrect, the objective is to use only Zipcode that are assign after optimization
            zipcode = correct_zip(str(zipcode))
            da_zipcode = correct_zip(str(w_opt['DA ZIP'][r]))
            carrier = w_opt['Carrier'][r]
            state = Zip_lat_long[da_zipcode][1]
            distance, Zip_lat_long, _ = compute_distance2(zipcode,da_zipcode,Zip_lat_long)
            flat = Pricing[state][carrier]['Flat']
            breakpoint = Pricing[state][carrier]['Break']
            extra = Pricing[state][carrier]['Extra']
            if distance <  breakpoint:
                lmcost=flat
            else :
                lmcost=flat+ (distance - breakpoint) * extra
            Optimized_LM.append(lmcost)
        else:
            Optimized_LM.append(0)
    
    
    w_zip['Estimated_Unit_Cost'] = Current_LM
    w_opt['Estimated_Unit_Cost'] = Optimized_LM
    w_zip['Estimated_LM_Cost'] = w_zip['Estimated_Unit_Cost'] * w_zip['Volume']
    w_opt['Estimated_LM_Cost'] = w_opt['Estimated_Unit_Cost'] * w_opt['Volume']
    Current_LM_Cost = w_zip['Estimated_LM_Cost'].sum()
    Optimized_LM_Cost = w_opt['Estimated_LM_Cost'].sum()
    
    # Computation of LH Cost
    w_lh = w_zip.groupby(['DA ZIP','Carrier'])['Volume'].sum()
    w_lh = w_lh.reset_index()
    w_lh['Weight_per_truck'] = w_lh['Volume']*weight_per_volume/nb_trucks
    w_lh_opt = w_opt.groupby(['DA ZIP','Carrier'])['Volume'].sum()
    w_lh_opt = w_lh_opt.reset_index()
    w_lh_opt['Weight_per_truck'] = w_lh_opt['Volume']*weight_per_volume/nb_trucks    
    Current_LH = []
    Optimized_LH = []
    
    for r in range(len(w_lh)):
        zipcode = correct_zip(str(w_lh['DA ZIP'][r]))
        state = Zip_lat_long[zipcode][1]
        weight = w_lh['Weight_per_truck'][r]
        try:
            for dfc_state in percDestin[state].keys():
                dfc_zip = DFC_Dict[dfc_state]['Zipcode']
                percentage = percDestin[state][dfc_state]
                distance, Zip_lat_long, _ = compute_distance2(zipcode,dfc_zip,Zip_lat_long)
                slope = coefficient["tot_mile_wt"]+coefficient["tot_mile_wt"]*distance
                intercept = coefficient['Intercept']+coefficient['tot_mile_cnt']*distance
                cost_opening = intercept + weight_treshold_ltl * slope
                Da_Dfc.setdefault(zipcode,{}).setdefault(dfc_state, {"distance":distance, "percentage" : percentage,"slope": slope,"cost_opening": cost_opening})
            slope = sum(Da_Dfc[zipcode][dfc_state]["slope"]*Da_Dfc[zipcode][dfc_state]["percentage"] for dfc_state in Da_Dfc[zipcode].keys())
            cost_opening = sum(Da_Dfc[zipcode][dfc_state]["cost_opening"]*Da_Dfc[zipcode][dfc_state]["percentage"] for dfc_state in Da_Dfc[zipcode].keys())
            Da_Dfc[zipcode]['Global']={'slope' : slope, 'cost_opening' : cost_opening}
        except KeyError:
    #        Just assume we take the previous cost
            Da_Dfc.setdefault(zipcode,{}).setdefault('Global',{'slope' : 0.1, 'cost_opening' : 70, 'Warning':"This Da doesn't have real slope or cost of opening"}  )  
        if weight < weight_treshold_ltl:
            lhcost = Da_Dfc[zipcode]['Global']['cost_opening']
        else : 
            lhcost = Da_Dfc[zipcode]['Global']['cost_opening'] + Da_Dfc[zipcode]['Global']['slope'] * (weight - weight_treshold_ltl) 
        lhcost = lhcost * nb_trucks
        
        Current_LH.append(lhcost)
    w_lh['LH Cost'] = Current_LH
    Current_LH_Cost = w_lh['LH Cost'].sum()
    
    
    for r in range(len(w_lh_opt)):
        zipcode = correct_zip(str(w_lh_opt['DA ZIP'][r]))
        state = Zip_lat_long[zipcode][1]
        weight = w_lh_opt['Weight_per_truck'][r]
        try:
            for dfc_state in percDestin[state].keys():
                dfc_zip = DFC_Dict[dfc_state]['Zipcode']
                percentage = percDestin[state][dfc_state]
                distance, Zip_lat_long, _ = compute_distance2(zipcode,dfc_zip,Zip_lat_long)
                slope = coefficient["tot_mile_wt"]+coefficient["tot_mile_wt"]*distance
                intercept = coefficient['Intercept']+coefficient['tot_mile_cnt']*distance
                cost_opening = intercept + weight_treshold_ltl * slope
                Da_Dfc.setdefault(zipcode,{}).setdefault(dfc_state, {"distance":distance, "percentage" : percentage,"slope": slope,"cost_opening": cost_opening})
            slope = sum(Da_Dfc[zipcode][dfc_state]["slope"]*Da_Dfc[zipcode][dfc_state]["percentage"] for dfc_state in Da_Dfc[zipcode].keys())
            cost_opening = sum(Da_Dfc[zipcode][dfc_state]["cost_opening"]*Da_Dfc[zipcode][dfc_state]["percentage"] for dfc_state in Da_Dfc[zipcode].keys())
            Da_Dfc[zipcode]['Global']={'slope' : slope, 'cost_opening' : cost_opening}
        except KeyError:
    #        Just assume we take the previous cost
            Da_Dfc.setdefault(zipcode,{}).setdefault('Global',{'slope' : 0.1, 'cost_opening' : 70, 'Warning':"This Da doesn't have real slope or cost of opening"}  )  
        if weight < weight_treshold_ltl:
            lhcost = Da_Dfc[zipcode]['Global']['cost_opening']
        else : 
            lhcost = Da_Dfc[zipcode]['Global']['cost_opening'] + Da_Dfc[zipcode]['Global']['slope'] * (weight - weight_treshold_ltl) 
        lhcost = lhcost * nb_trucks
        
        Optimized_LH.append(lhcost)
    w_lh_opt['LH Cost'] = Optimized_LH
    Optimized_LH_Cost = w_lh_opt['LH Cost'].sum()
    
    
    
    Current_Total_Cost = Current_LH_Cost + Current_LM_Cost
    Optimized_Total_Cost = Optimized_LH_Cost + Optimized_LM_Cost

    Current_Cost.append(Current_Total_Cost)
    Optimized_Cost.append(Optimized_Total_Cost)

df = pd.DataFrame({'Optimized': Optimized_Cost,'Current': Current_Cost})
# Fix column names
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\Output.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()