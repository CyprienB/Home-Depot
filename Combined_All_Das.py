# -*- coding: utf-8 -*-
"""
Created on Fri Sep 29 14:53:36 2017
This file combine all other Python files with some improvments to avoid useless processes
This is an optimization of the Home Depot .com Delivery network, with objective function being the total cost (line haul + last Mile) 
@author: Cyprien Bastide, Steven (Gao) Ming, Edson David Silva Moreno
"""



#Import the differents Modules that are going to be used in this file
import sys
#   This Module is used to open excel files
import openpyxl as xl
from statsmodels.formula.api import ols
# Neig_states return the neighboring state of the input state
# is an easier way to call cell inan excel file
# instance return the number of lines in an excel spreadsheet
# Compute distance2 compute the distances from 2 zip code and the lat long database
# Correct zip add 0 in front of postal codes that are not 5 digits long
# Get last mile pricing returns a dictionnary containing the info to compute the last mile cost
from Procedures import neig_states, cell, instance, compute_distance2, correct_zip, get_lm_pricing, averageOrig
#tqdm is used to create progress bar, however, if computation is too fast it might cause bugs
from tqdm import tqdm
# Pulp is the optimization engine
import pulp
#time allow to compute time elapsed for some taks
import time
import pandas as pd
number_days = 30*6
weight_treshold_ltl = 200
nb_trucks = round(number_days*5/7)
weight_per_volume = 100
#coefficient= {'intercept':-3.683,'weight':0.1498,'dist':0.0537,'weight_dist':0.0001,'CA':0,"GA":-8.4855,"MD":-7.5867,"OH":3.4399}

#Extract Data from Excel
ltl_price = pd.read_excel('C:\HomeDepot_Excel_Files\Standard_File.xlsx', sheetname='ltl_price')
ltl_price.head()
ltl_price = ltl_price[(ltl_price['tot_shp_wt'] >= 200) & (ltl_price['tot_shp_wt'] <= 4999) & (ltl_price['aprv_amt'] <=1000)]
#ltl_price["tot_mile_cnt1"] = ltl_price["tot_mile_cnt"] - np.mean(ltl_price["tot_mile_cnt"])
#ltl_price["tot_shp_wt1"] = ltl_price["tot_shp_wt"] - np.mean(ltl_price["tot_shp_wt"])
ltl_price['tot_mile_wt'] = ltl_price['tot_mile_cnt'] * ltl_price['tot_shp_wt']
#ltl_price['orig_state'] = ltl_price["orig_state"].astype('category')
orig_state =pd.get_dummies(ltl_price['orig_state'])
full_data = pd.concat([ltl_price,orig_state], axis=1)      

# fit our model with .fit() and show results
linehaul_model = ols('aprv_amt ~ tot_mile_cnt + tot_shp_wt + + tot_mile_wt + CA + GA + OH + MD', data=full_data).fit()
# summarize our model
linehaul_model_summary = linehaul_model.summary()
print(linehaul_model_summary)

#Terms & Coeff
variables = [linehaul_model.params.index.tolist()][0]
# Filter and Rename orig_state
for i in range(0, len(variables)):
    if variables[i].find("orig_state") != -1:
        variables[i] = variables[i][-3]+variables[i][-2]

coeff = [linehaul_model.params.tolist()][0]
# Convert two lists to a dictionary
coefficient = dict(zip(variables,coeff))
"""
###############################################################
###############################################################

This part upload the information from the excel spreadsheet

###############################################################
###############################################################
"""
print('Import Workbook') 
# Open Worksheet that contains list of DA, List of Zip code with Volume, Pricing spreadsheet,...
wb = xl.load_workbook('C:\HomeDepot_Excel_Files\Standard_File.xlsx')
# Open All Different Spreadsheet
w_neig = wb['List_of_Neighboring_States']
w_da = wb['DA_List']
w_zip = wb['Zip_Allocation_and_Pricing']
w_lm_pricing = wb["LM_Pricing"]
w_range = wb["Zip_Range"]
w_dfc = wb["DFC list"]
# Import Database of Zipcode Latitude and Longitude
print('Open Lat Long Database')
wdata = xl.load_workbook('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
wslatlong = wdata['Zip']

#Importing Excel sheet as Panda Data Frames to create Dictionary with every destination state as a Key and each Key has a nested Dictionary with the weight (percentage) of invoices coming from every origin for LTL pricing.
print('Import Database LTL')
wbLtl = pd.ExcelFile('C:\HomeDepot_Excel_Files\Standard_File.xlsx')
ltl_price = wbLtl.parse('ltl_price', converters={'dest_zip': str,'orig_zip': str})



"""
###############################################################
###############################################################

This part create the different dictionnaries that are going to be used to solve the optimization problem
Multiple dictionnaries are going to be created to represent DAs and Zipcode because we will need different keys to call them

###############################################################
###############################################################
"""

print ("Create all Dictionnaries")
# Get number of DA and of Zip
n_da= instance(w_da)
n_zip = instance(w_zip)
linelatlong = instance(wslatlong)
# Create dictionnaries for DA and Zip(with volume in tuple) Grouped by State, 
# This is useful because Arcs are going to be created based on neighgboring states
# Since distances don't depend on Carrier, if multiple DA are in the same zipcode only one will be counted

# Dictionnary for DA  {State : [Da_ZipCode]} 
State_Da_dict = {}
for r in range(n_da):
    state = cell(w_da,r+2,3) 
    da_zip = correct_zip(str(cell(w_da,r+2,2))) 
#        Remove duplicate: DA in same zipcode but different carriers
    State_Da_dict[state] = list(set().union(State_Da_dict.setdefault(state,[]),[da_zip]))


# Dictionnary for Zip {State : [( zipcode, volume)]}
State_Zip_dict = {}
for r in range(n_zip):
    state = cell(w_zip,r+2,2)
    zipcode = correct_zip(str(cell(w_zip,r+2,1))) 
    volume = cell(w_zip,r+2,7)
    
    State_Zip_dict.setdefault(state, []).append((zipcode, volume))

#Create Dictionnary for Da { Zip:(Zip, State, [Carrier])} will be useful o compute distances

DA_ZipCode_Dict = {}
for r in range(n_da):
    zipcode = correct_zip(str(cell(w_da,r+2,2)))
    state = cell(w_da, r+2, 3)
    carrier = cell(w_da,2+r, 4)
    
    DA_ZipCode_Dict.setdefault(zipcode,{'Zip':zipcode, 'State':state, 'Carrier':[carrier]}) 

    DA_ZipCode_Dict[zipcode]['Carrier'] = list(set().union(DA_ZipCode_Dict[zipcode]['Carrier'],[carrier]))



# Other dictionnary for Da, { Zip + Carrier : (Zip, State, Carrier)}

DAC_ZipCode_Dict = {}
for r in range(n_da):
    zipcode = correct_zip(str(cell(w_da,r+2,2)))
    carrier = cell(w_da, r+2, 4)
    state = cell(w_da, r+2, 3)
    DAC_ZipCode_Dict[zipcode+' '+ carrier] = {'Zip':zipcode, 'State':state, 'Carrier':carrier}


#Create Dictionnary for Zipcode (Zip, Volume ,State)
ZipCode_Dict={}
for r in range(n_zip):
    zipcode = correct_zip(str(cell(w_zip,2+r, 1)))
    volume = cell(w_zip, r+2, 7)
    state = cell(w_zip,r+2, 2)
    ZipCode_Dict[zipcode]={'Zip':zipcode,'Volume':volume, 'State':state}         

# Create LM Pricing Dictionnary
Pricing=get_lm_pricing(w_lm_pricing)

# Get arc max range Dictionnary {State : Max range}
Range = { cell(w_range,r+2,2) : cell(w_range,r+2,3) for r in range(instance(w_range))}


# Create dictionnary for the database lat long{Zip : (lat,long)}
Zip_lat_long = {}
for r in range(linelatlong):
    zipcode = correct_zip(str(cell(wslatlong,r+2,1)))
    lat = cell(wslatlong,r+2,2)
    long = cell(wslatlong,r+2,3)
    Zip_lat_long[zipcode] = (lat,long)

#Create dictionary with State Destination as a Key and nested dictionary with weight by origin 
percDestin = averageOrig(ltl_price)

# Create Dictionnary of DFC {State : {Name, Zip, State}}
DFC_Dict={}
nbdfc = instance(w_dfc)
for r in range(nbdfc):
    name = cell(w_dfc,r+2,1)
    zipcode = correct_zip(str(cell(w_dfc,r+2,2)))
    state = cell(w_dfc,r+2,3)
    DFC_Dict[state]={'State':state,'Name':name,'Zipcode':zipcode}
"""
###############################################################
###############################################################

This part compute the distance between Das and DFC and the pricing based on parameters and distance

\\ Problem for new DAs in States without Das currently (Key Error)

###############################################################
############################################################### 
"""
# Relation Da_Dfc is in format {da : {state_dfc : {distance, percentage from state}, global : {slope,cost_opening}}
a = 0
Da_Dfc = {}

for da in DA_ZipCode_Dict.keys():
    
    da_state = DA_ZipCode_Dict[da]["State"]
    try:
        for dfc_state in percDestin[da_state].keys():
            
            dfc_zip = DFC_Dict[dfc_state]['Zipcode']
            percentage = percDestin[da_state][dfc_state]
            
            distance, Zip_lat_long, b = compute_distance2(da,dfc_zip,Zip_lat_long)
            
            slope = coefficient["tot_mile_wt"]+coefficient["tot_mile_wt"]*distance
            
            intercept = coefficient['Intercept']+coefficient[dfc_state]+coefficient['tot_mile_cnt']*distance
            
            cost_opening = intercept + weight_treshold_ltl * slope
            
            Da_Dfc.setdefault(da,{}).setdefault(dfc_state, {"distance":distance, "percentage" : percentage,"slope": slope,"cost_opening": cost_opening})
            if b == 1 :
                a += 1
        slope = sum(Da_Dfc[da][dfc_state]["slope"]*Da_Dfc[da][dfc_state]["percentage"] for dfc_state in Da_Dfc[da].keys())
        
        cost_opening = sum(Da_Dfc[da][dfc_state]["cost_opening"]*Da_Dfc[da][dfc_state]["percentage"] for dfc_state in Da_Dfc[da].keys())
        
        Da_Dfc[da]['Global']={'slope' : slope, 'cost_opening' : cost_opening}
        
    except KeyError:
#        Just assume we take the previous cost
        Da_Dfc.setdefault(da,{}).setdefault('Global',{'slope' : slope, 'cost_opening' : cost_opening, 'Warning':"This Da doesn't have real slope or cost of opening"}  )  

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
# Create combination if Volume is not 0, have list of all combination [[da,zip,distance]]
    for da in da_list:
        for pc in zip_list:
            zipcode = pc[0]
            volume = pc[1]
            distance, Zip_lat_long, b = compute_distance2(da, zipcode, Zip_lat_long)
            combination.append([da,zipcode,distance])
            if b == 1 :
                a += 1


# Update if new zipcodes have been added
if a != 0:
    print("Update Database")
    ZipList = Zip_lat_long.keys()
    c = 0
    for r in ZipList:
        wslatlong.cell(row = c+2,column = 1).value = r
        wslatlong.cell(row = c+2,column = 2).value = Zip_lat_long[r][0]
        wslatlong.cell(row = c+2,column = 3).value = Zip_lat_long[r][1]
        c+=1
    wdata.save('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
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
for i in tqdm(range(len(combination))):
    da = combination[i][0]
    zipcode = combination[i][1] 
    distance = combination[i][2]
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
        if distance > 75 : # Opportunity cost
            lmcost += 0
        Arcs[pc][da]['lm_cost'] = lmcost

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
    return Arcs[pc][da]['lm_cost']*ZipCode_Dict[pc]['Volume']*Arcs[pc][da]['variable']
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
    prob += pulp.lpSum(Zip_temp)-1500*DAC_ZipCode_Dict[da]['opening_variable'] <= 0
    
 # Constraint over the weight variable   
for da in DAC_ZipCode_Dict.keys():
    prob += DAC_ZipCode_Dict[da]["Weight_variable"] >= 0
    prob += DAC_ZipCode_Dict[da]["Weight_variable"] >= pulp.lpSum([ZipCode_Dict[pc]['Volume']*Arcs[pc][da]['variable'] for pc in Arcs.keys() if da in Arcs[pc]]) / nb_trucks * weight_per_volume -weight_treshold_ltl 

# Open a certain number of DA
#    
#prob += pulp.lpSum([DAC_ZipCode_Dict[da]['opening_variable'] for da in DAC_ZipCode_Dict.keys() ])>= 120

# The problem is solved using PuLP's choice of Solver
print("Solve Problem")
start_time = time.clock()
solve= pulp.solvers.GUROBI(timeLimit = 300)

solve.actualSolve(prob)
end_time = time.clock()
print(end_time-start_time)

# The status of the solution is printed to the screen
print("Status:", pulp.LpStatus[prob.status])


# The optimised objective function value is printed to the screen    
print ("total cost", pulp.value(prob.objective))

#Create workbook for results
w_result = xl.Workbook()
wresultassign = w_result.create_sheet('Optimization Results Assignment')
wresultda =  w_result.create_sheet('Optimization Results DA')
# export results on excel

print("Exporting Results")

wresultassign.cell(row=1,column=1).value= "ZipCode"
wresultassign.cell(row=1,column=2).value= "Carrier"
wresultassign.cell(row=1,column=3).value= "DaZipCode"
wresultassign.cell(row=1,column=4).value= 'DA and Carrier'
wresultassign.cell(row=1,column=5).value= 'Volume'
wresultassign.cell(row=1,column=6).value= 'Unit Cost'
wresultassign.cell(row=1,column=7).value= 'Total Cost'
wresultassign.cell(row=1,column=8).value= 'Assignment Variable'
# Print Results on excel
r=2
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        if Arcs[pc][da]['variable'].varValue > 0.01 :
            wresultassign.cell(row=r,column=1).value= pc
            wresultassign.cell(row=r,column=2).value= da[6:]
            wresultassign.cell(row=r,column=3).value= da[:5]
            wresultassign.cell(row=r,column=4).value= da
            wresultassign.cell(row=r,column=5).value= ZipCode_Dict[pc]['Volume']            
            wresultassign.cell(row=r,column=6).value= Arcs[pc][da]['lm_cost']
            wresultassign.cell(row=r,column=7).value= Arcs[pc][da]['lm_cost'] * Arcs[pc][da]['variable'].varValue*ZipCode_Dict[pc]['Volume']  
            wresultassign.cell(row=r,column=8).value= Arcs[pc][da]['variable'].varValue
            r+=1
            
            
wresultda.cell(row=1,column=1).value= "Da"
wresultda.cell(row=1,column=2).value= "Carrier"
wresultda.cell(row=1,column=3).value= "Da zip"
wresultda.cell(row=1,column=4).value= 'Volume above 200'
wresultda.cell(row=1,column=5).value= 'lh cost'

# Print Results on excel
r=2


Useful_Da = []
for da in DAC_ZipCode_Dict.keys():
    if DAC_ZipCode_Dict[da]['opening_variable'].varValue == 1:
        wresultda.cell(row=r,column=1).value= da
        wresultda.cell(row=r,column=2).value= da[6:]
        wresultda.cell(row=r,column=3).value= da[:5]
        wresultda.cell(row=r,column=4).value= DAC_ZipCode_Dict[da]["Weight_variable"].varValue
        wresultda.cell(row=r,column=5).value= (DAC_ZipCode_Dict[da]["Weight_variable"].varValue * Da_Dfc[da[:5]]['Global']['slope']+ Da_Dfc[da[:5]]['Global']['cost_opening'])*nb_trucks
        wresultda.cell(row=r,column=6).value= DAC_ZipCode_Dict[da]["opening_variable"].varValue     
        
 # Return List of useful DA       
        
        Useful_Da = list(set().union([da],Useful_Da))
        
        r+=1

print("Number of Useful Das:", len(Useful_Da))


print("Save File")

w_result.save("C:\HomeDepot_Excel_Files\Optimized.xlsx")
         