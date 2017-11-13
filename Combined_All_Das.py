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
from tqdm import tqdm
# This module is the optimization engine
import pulp
# This module is used to compute time elapsed for a task
import time
# This module is used to convert spreadsheet into a specific dataframe for analysis
import pandas as pd
# neig_states: it returns the neighboring state of the input state
# compute_distance2: it computes the distances from 2 zip code and the lat long database
# correct_zip: it adds 0 in front of postal codes that are not 5 digits long
# get_lm_pricing: it returns a dictionnary containing the info to compute the last mile cost
# averageOrig: # it returns the dictionary of every State Destination with weighted origin 
from Procedures import neig_states, compute_distance2, correct_zip, get_lm_pricing, averageOrig, geocode2

optimization_time = 500
oportunity_threshold = 50
oportunity_cost = 10
number_days = 30*6
weight_treshold_ltl = 200
nb_trucks = round(number_days*5/7)
weight_per_volume = 100
DA_to_DA_min_distance = 40*1.54

# Import and convert spreadsheets into panda dataframe
wb = pd.ExcelFile('C:\HomeDepot_Excel_Files\Standard_File.xlsx')
wd = pd.ExcelFile('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
w_neig = wb.parse('List_of_Neighboring_States')
w_da = wb.parse('DA_List')
w_zip = wb.parse('Zip_Allocation_and_Pricing')
w_range = wb.parse("Zip_Range")
w_dfc = wb.parse("DFC list")
wslatlong = wd.parse('Zip_Lat_Long')
ltl_price = wb.parse('ltl_price', converters={'dest_zip': str,'orig_zip': str})



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

# Create Dictionnary of DFC {State : {Name, Zip, State}}
DFC_Dict={}
for r in range(nbdfc):
    name = w_dfc['DFC'][r]
    zipcode = correct_zip(str(w_dfc['DFC ZIP'][r]))
    state = w_dfc['DFC State'][r]
    DFC_Dict[state]={'State':state,'Name':name,'Zipcode':zipcode}

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
###############################################################
###############################################################

This part create the combination DA Zipcode based on neighboring states and then compute the distances 

###############################################################
###############################################################
"""

# Create couple that need to have distance assigned to it based Neighboring states, 
# Go through every state, look at the Da inside, and assign them to all zip code in neighboring states

combination = []
combinationDA = []
# Iterate through the states
print('Creating combination Da Zipcode')
for state in tqdm(State_Da_dict.keys()):
#    Create list of Da in the state
    da_list = State_Da_dict[state]

#   Create list of Zipcode that are in neigh state
    zip_list = []
    neighboring_states= neig_states(state,w_neig)
    for n_state in neighboring_states:
        zip_list += State_Zip_dict[n_state]
# Create combination , have list of all combination [[da,zip,distance]]
    for da in da_list:
        for pc in zip_list:
            zipcode = pc[0]
            distance, Zip_lat_long, _ = compute_distance2(da, zipcode, Zip_lat_long)
            combination.append([da,zipcode,distance])


# Iterate through the states
print('Creating combination Da_Da')
for state in tqdm(State_DAC_dict.keys()):
#    Create list of Da in the state
    da_list = State_DAC_dict[state]

#   Create list of Zipcode that are in neigh state
    to_da_list = []
    neighboring_states= neig_states(state,w_neig)
    for n_state in neighboring_states :
        if n_state in State_DAC_dict.keys():
            to_da_list+= State_DAC_dict[n_state]
# Create combination , have list of all combination [[da,zip,distance]]
    for da in da_list:
        for to_da in to_da_list:
            if da != to_da:
                zipda = da[:5]  
                zipcode = to_da[:5]
                distance, Zip_lat_long, _ = compute_distance2(zipda, zipcode, Zip_lat_long)
                combinationDA.append([da,to_da,distance])


# Update if new zipcodes 
print("Update Database")
ZipList= []
LatList = []
LongList = []
StateList = []
for zipcode in Zip_lat_long.keys():
    latitude = Zip_lat_long[zipcode][0][0]
    longitude = Zip_lat_long[zipcode][0][1]
    state = Zip_lat_long[zipcode][1]
    ZipList.append(zipcode)
    LatList.append(latitude)
    LongList.append(longitude)
    StateList.append(state)

Database = pd.DataFrame({'ZipCode':ZipList, 'Latitude' : LatList, 'Longitude': LongList, 'State':StateList})
Database = Database[['ZipCode','Latitude','Longitude', 'State']]
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx', engine='xlsxwriter')
Database.to_excel(writer,sheet_name = 'Zip_Lat_Long', index = False)
writer.save()
print('Database updated')
    
    
  
"""
###############################################################
###############################################################

This part is the HDU optimization model (includes last mile and line haul)
and uses all DA(based on previous optimization)

\For now arcs are recreated while we could use previous dictionnary
\Model only has one treshold
###############################################################
###############################################################
"""
# Create dictionnary for the arcs
Arcs={}
for da, zipcode, distance in tqdm(combination):

#        Create an arc only if distance between DA and Zip is less than the Zip's state threshold in this model we only use volume above zero
    if distance< Range[ZipCode_Dict[zipcode]['State']] and ZipCode_Dict[zipcode]['Volume']>0:
        for carrier in DA_ZipCode_Dict[da]['Carrier']:
            Arcs.setdefault(zipcode,{}).setdefault(da +" "+carrier,{'distance' : distance})
                
                
# Compute Costs for the arcs
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        
        distance = Arcs[pc][da]['distance']        
        da_state = DAC_ZipCode_Dict[da]['State']
        da_carrier = DAC_ZipCode_Dict[da]['Carrier']
        
        try:
            flat = Pricing[da_state][da_carrier]['Flat']
            breakpoint = Pricing[da_state][da_carrier]['Break']
            extra = Pricing[da_state][da_carrier]['Extra']
        except KeyError:
            sys.exit(("LM_Cost spreadsheet does not contain pricing info for couple state-carrier %s, %s, need to update Standard File" %(da_state,da_carrier)))
            
#        Check if distance Da_Zip is within flat distance
        if distance <  breakpoint:
            lmcost=flat
        else :
            lmcost=flat+ (distance - breakpoint) * extra
        if distance > oportunity_threshold : # Oportunity cost
            lm_oport_cost = lmcost + oportunity_cost
        else:
            lm_oport_cost = lmcost
        Arcs[pc][da]['lm_cost'] = lmcost
        Arcs[pc][da]['lm_oport_cost'] = lm_oport_cost

#   { Zip : {DA+Carrier :{distance, lm_cost (come in next step), var(come in two steps)}}
print("Create Model")

# Create Model
prob = pulp.LpProblem("Minimize HDU Cost",pulp.LpMinimize)

# Design arcs
for pc in Arcs.keys():
    for da in Arcs[pc].keys():

        var = pulp.LpVariable("Arc_%s_%s)" % (pc,da),0,1,pulp.LpContinuous)
        Arcs[pc][da]['variable']=var
            
# Create variable for Das (OPEN AND VOLUME OVER 200)
for da in DAC_ZipCode_Dict.keys():

    ovar = pulp.LpVariable("Da_%s" % (str(da)),0,1,pulp.LpBinary)
    wvar = pulp.LpVariable("Da_%s_above_200LBS" % (str(da)))
    DAC_ZipCode_Dict[da]['opening_variable']=ovar
    DAC_ZipCode_Dict[da]['Weight_variable']=wvar


# Create Objective function : minimize cost
print("Create objective and Constraint")
def lmcost(pc,da):
    return (Arcs[pc][da]['lm_oport_cost']*ZipCode_Dict[pc]['Volume']+0.01*Arcs[pc][da]['distance'])*Arcs[pc][da]['variable']
def lhcost(da):
    zip_da = da[:5]
    return Da_Dfc[zip_da]['Global']['cost_opening']*DAC_ZipCode_Dict[da]['opening_variable'] + Da_Dfc[zip_da]['Global']['slope'] * DAC_ZipCode_Dict[da]["Weight_variable"]

prob += pulp.lpSum([lmcost(pc,da) for pc in Arcs.keys() for da in Arcs[pc].keys()]) + nb_trucks*pulp.lpSum([lhcost(da) for da in DAC_ZipCode_Dict.keys()])

# Create Constraint : every Zip is allocated
print("Create contraint 'every zipcode is assigned to a DA'")
for pc in tqdm(Arcs.keys()):          
    prob += pulp.lpSum([Arcs[pc][da]['variable'] for da in Arcs[pc].keys()]) == 1

# Volume only if DA open, limit the max number of Zip
for da in DAC_ZipCode_Dict.keys():
    Zip_temp = []
    for pc in Arcs.keys():
        try:
            Zip_temp.append(Arcs[pc][da]['variable'])
        except:
#            Whatever lline to prevent error in case da is not in zip dict
            j=10      
    prob += pulp.lpSum(Zip_temp)-3000*DAC_ZipCode_Dict[da]['opening_variable'] <= 0
    
 # Constraint over the weight variable   
for da in DAC_ZipCode_Dict.keys():
    prob += DAC_ZipCode_Dict[da]["Weight_variable"] >= 0
    prob += DAC_ZipCode_Dict[da]["Weight_variable"] >= pulp.lpSum([ZipCode_Dict[pc]['Volume']*Arcs[pc][da]['variable'] for pc in Arcs.keys() if da in Arcs[pc]]) / nb_trucks * weight_per_volume -weight_treshold_ltl 

# Open a certain number of DA
#    
#prob += pulp.lpSum([DAC_ZipCode_Dict[da]['opening_variable'] for da in DAC_ZipCode_Dict.keys() ])>= 120



# Constrain distance between das

for line in combinationDA:
    da1 = line[0]
    da2 = line[1]
    distance = line[2]
    carrier = da1[6:]
    state = DAC_ZipCode_Dict[da1]['State']
#    if distance < Pricing[state][carrier]['Break']:
#        prob += DAC_ZipCode_Dict[da1]['opening_variable']+DAC_ZipCode_Dict[da2]['opening_variable'] <= 1 
    if distance < DA_to_DA_min_distance:
        prob += DAC_ZipCode_Dict[da1]['opening_variable']+DAC_ZipCode_Dict[da2]['opening_variable'] <= 1
 
    
 # The problem is solved using PuLP's choice of Solver
print("Solve Problem")

solve= pulp.solvers.GUROBI(timeLimit = optimization_time)

solve.actualSolve(prob)


# The status of the solution is printed to the screen
print("Status:", pulp.LpStatus[prob.status])


# The optimised objective function value is printed to the screen    
print ("total cost", pulp.value(prob.objective))

"""
#########################
Put results into dataframe
##########################
"""

print("Exporting Results")

column_names=["ZipCode",
         "Carrier",
         "DaZipCode",
         'DA and Carrier',
         'Volume',
         'Unit Cost',
         'Total Cost',
         'Oportunity_cost',
         'Assignment Variable',
         "distance"]
# Print Results as DataFrame
Assign_Results = []
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        if Arcs[pc][da]['variable'].varValue > 0.001 :
            Assign_Results.append([pc,
                                   da[6:],
                                   da[:5],
                                   da,
                                   ZipCode_Dict[pc]['Volume'],
                                   Arcs[pc][da]['lm_cost'],
                                   Arcs[pc][da]['lm_cost'] * Arcs[pc][da]['variable'].varValue*ZipCode_Dict[pc]['Volume'],
                                   Arcs[pc][da]['lm_oport_cost'],
                                   Arcs[pc][da]['variable'].varValue,
                                   Arcs[pc][da]['distance']])
Assign_Results = pd.DataFrame(Assign_Results,columns = column_names)
            
            
column_names_da= ["Da",
              "Carrier",
              "Da zip",
              'Volume above 200',
              'lh cost',
              'Opening_variable']

DA_Results = []
Useful_Da = []
for da in DAC_ZipCode_Dict.keys():
    if DAC_ZipCode_Dict[da]['opening_variable'].varValue > 0.0001:
        DA_Results.append([da,
                           da[6:],
                           da[:5],
                           DAC_ZipCode_Dict[da]["Weight_variable"].varValue,
                           (DAC_ZipCode_Dict[da]["Weight_variable"].varValue * Da_Dfc[da[:5]]['Global']['slope']+ Da_Dfc[da[:5]]['Global']['cost_opening'])*nb_trucks,
                            DAC_ZipCode_Dict[da]["opening_variable"].varValue])
 # Return List of useful DA       
        Useful_Da = list(set().union([da],Useful_Da))
        
        r+=1
DA_Results = pd.DataFrame(DA_Results,columns = column_names_da)


print("Number of Useful Das:", len(Useful_Da))


#print("Save File")
#
#w_result.save("C:\HomeDepot_Excel_Files\Optimized.xlsx")
         


"""
#####################################
#####################################
Optimize 0 Volume Postal and leftovers
#####################################
#####################################
"""
#
# Create dictionnary for the arcs
Arcs0={}
for i in tqdm(range(len(combination))):
    da = combination[i][0]
    zipcode = combination[i][1] 
    distance = combination[i][2]

    if zipcode not in Arcs.keys():
        for carrier in DA_ZipCode_Dict[da]['Carrier']:
            if da + " " + carrier in Useful_Da:
                Arcs0.setdefault(zipcode,{}).setdefault(da +" "+carrier,{'distance' : distance})
                
                
# Compute Costs for the arcs
for pc in Arcs0.keys():
    for da in Arcs0[pc].keys():
        
        distance = Arcs0[pc][da]['distance']        
        da_state = DAC_ZipCode_Dict[da]['State']
        da_carrier = DAC_ZipCode_Dict[da]['Carrier']
        
        try:
            flat = Pricing[da_state][da_carrier]['Flat']
            breakpoint = Pricing[da_state][da_carrier]['Break']
            extra = Pricing[da_state][da_carrier]['Extra']
        except KeyError:
            sys.exit(("LM_Cost spreadsheet does not contain pricing info for couple state-carrier %s, %s, need to update Standard File" %(da_state,da_carrier)))
            
#        Check if distance Da_Zip is within flat distance
        if distance <  breakpoint:
            lmcost=flat
        else :
            lmcost=flat+ (distance - breakpoint) * extra
        if distance > oportunity_threshold : # Oportunity cost
            lm_oport_cost = lmcost + oportunity_cost
        else:
            lm_oport_cost = lmcost
        Arcs0[pc][da]['lm_cost'] = lmcost
        Arcs0[pc][da]['lm_oport_cost'] = lm_oport_cost

#   { Zip : {DA+Carrier :{distance, lm_cost (come in next step), var(come in two steps)}}
print("Create Model")

# Create Model
prob = pulp.LpProblem("Minimize LM Cost",pulp.LpMinimize)

# Design arcs
for pc in Arcs0.keys():
    for da in Arcs0[pc].keys():

        var = pulp.LpVariable("Arc_%s_%s)" % (pc,da),0,1,pulp.LpContinuous)
        Arcs0[pc][da]['variable']=var
            

# Create Objective function : minimize cost
print("Create objective and Constraint")
def lmcost2(pc,da):
    return (Arcs0[pc][da]['lm_oport_cost']+0.005*Arcs0[pc][da]['distance'])*Arcs0[pc][da]['variable']

prob += pulp.lpSum([lmcost2(pc,da) for pc in Arcs0.keys() for da in Arcs0[pc].keys()])

# Create Constraint : every Zip is allocated
print("Create contraint 'every zipcode is assigned to a DA'")
for pc in tqdm(Arcs0.keys()):          
    prob += pulp.lpSum([Arcs0[pc][da]['variable'] for da in Arcs0[pc].keys()]) == 1

# The problem is solved using PuLP's choice of Solver
print("Solve Problem")

solve= pulp.solvers.GUROBI(timeLimit = optimization_time)

solve.actualSolve(prob)

"""
#########################
Put results into dataframe
##########################
"""

print("Exporting Results")

column_names=["ZipCode",
         "Carrier",
         "DaZipCode",
         'DA and Carrier',
         'Volume',
         'Unit Cost',
         'Total Cost',
         'Oportunity_cost',
         'Assignment Variable',
         "distance"]
# Print Results as DataFrame
Assign_Results2 = []
for pc in Arcs0.keys():
    for da in Arcs0[pc].keys():
        if Arcs0[pc][da]['variable'].varValue > 0.001 :
            Assign_Results2.append([pc,
                                   da[6:],
                                   da[:5],
                                   da,
                                   ZipCode_Dict[pc]['Volume'],
                                   Arcs0[pc][da]['lm_cost'],
                                   Arcs0[pc][da]['lm_cost'] * Arcs0[pc][da]['variable'].varValue*ZipCode_Dict[pc]['Volume'],
                                   Arcs0[pc][da]['lm_oport_cost'],
                                   Arcs0[pc][da]['variable'].varValue,
                                   Arcs0[pc][da]['distance']])
Assign_Results = Assign_Results.append(pd.DataFrame(Assign_Results2,columns = column_names), ignore_index=True)
            
print("Write Excel")
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\Optimized_oportunity.xlsx', engine='xlsxwriter')
Assign_Results.to_excel(writer,'AssignmentResults', index = False)
DA_Results.to_excel(writer,'OptimizedDA', index = False)
writer.save()


"""
##############################
##############################
Cost Analysis
##############################
##############################
"""

Optimized_LM_Cost = Assign_Results['Total Cost'].sum()
Optimized_LH_Cost = DA_Results['lh cost'].sum()
Total_Optimized_Cost = Optimized_LH_Cost+Optimized_LM_Cost


# Computation of LM Cost of Current Network
New_Column = []
for r in range(len(w_zip)):
    zipcode = correct_zip(str(w_zip['Zip#'][r]))
    if zipcode in Assign_Results['ZipCode'].values:  # Maybe incorrect, the objective is to use only Zipcode that are assign after optimization
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
        New_Column.append(lmcost)
    else:
        New_Column.append(0)
        
w_zip['Estimated_Unit_Cost'] = New_Column
w_zip['Estimated_LM_Cost'] = w_zip['Estimated_Unit_Cost'] * w_zip['Volume']
w_zip['Difference'] = w_zip['Estimated_LM_Cost']- w_zip['Cost']
Current_LM_Cost = w_zip['Estimated_LM_Cost'].sum()

# Computation of LH Cost
w_lh = w_zip.groupby(['DA ZIP','Carrier'])['Volume'].sum()
w_lh = w_lh.reset_index()
w_lh['Weight_per_truck'] = w_lh['Volume']*weight_per_volume/nb_trucks

New_Column = []

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
    
    New_Column.append(lhcost)

w_lh['LH Cost'] = New_Column


Current_LH_Cost = w_lh['LH Cost'].sum()

Current_Total_Cost = Current_LH_Cost + Current_LM_Cost
percentage_savings = (Current_Total_Cost-Total_Optimized_Cost)/Current_Total_Cost*100
print('Number of Das :', len(Useful_Da))
print('Current LM Cost :', Current_LM_Cost, ', Optimized Network LM Cost :', Optimized_LM_Cost)
print('Current LH Cost :', Current_LH_Cost, ', Optimized Network LH Cost :', Optimized_LH_Cost)
print('Current Total Cost :', Current_Total_Cost, ', Optimized Network Total Cost :', Total_Optimized_Cost)
print('Savings in percentage :', percentage_savings)