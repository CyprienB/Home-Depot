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
# Neig_states return the neighboring state of the input state
# cell is an easier way to call cell inan excel file
# instance return the number of lines in an excel spreadsheet
# Compute distance2 compute the distances from 2 zip code and the lat long database
# Correct zip add 0 in front of postal codes that are not 5 digits long
# Get last mile pricing returns a dictionnary containing the info to compute the last mile cost
from Procedures import neig_states, cell, instance, compute_distance2, correct_zip, get_lm_pricing
#tqdm is used to create progress bar, however, if computation is too fast it might cause bugs
from tqdm import tqdm
# Pulp is the optimization engine
import pulp
#time allow to compute time elapsed for some taks
import time


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

# Import Database of Zipcode Latitude and Longitude
print('Open Database')
wdata = xl.load_workbook('C:\HomeDepot_Excel_Files\Zip_latlong.xlsx')
wslatlong = wdata['Zip']

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
    try:
#        Remove duplicate: DA in same zipcode but different carriers
        State_Da_dict[state] = list(set().union(State_Da_dict[state],[da_zip]))
    except KeyError:
        State_Da_dict[state]=[]
        State_Da_dict[state].append(da_zip)

# Dictionnary for Zip {State : [( zipcode, volume)]}
State_Zip_dict = {}
for r in range(n_zip):
    state = cell(w_zip,r+2,2)
    zipcode = correct_zip(str(cell(w_zip,r+2,1))) 
    volume = cell(w_zip,r+2,7)
    try:
        State_Zip_dict[state].append((zipcode, volume))
    except KeyError:
        State_Zip_dict[state]=[]
        State_Zip_dict[state].append((zipcode, volume))
        
#Create Dictionnary for Da { Zip:(Zip, State, [Carrier])}

DA_ZipCode_Dict = {}
for r in range(n_da):
    zipcode = correct_zip(str(cell(w_da,r+2,2)))
    state = cell(w_da, r+2, 3)
    carrier = cell(w_da,2+r, 4)
    try :
        DA_ZipCode_Dict[zipcode]['Carrier'] = list(set().union(DA_ZipCode_Dict[zipcode]['Carrier'],[carrier]))
    except KeyError:
        DA_ZipCode_Dict[zipcode] = {'Zip':zipcode, 'State':state, 'Carrier':[carrier]}
       
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
# Create combination if Volume is not 0, have list of all combination [[da,zip]]
    for da in da_list:
        for pc in zip_list:
            zipcode = pc[0]
            volume = pc[1]
            if volume != 0:
                combination.append([da,zipcode])
                
# Compute distances for each combination
nb_distances = len(combination)
a = 0
print("Compute distances")
for i in tqdm(range(nb_distances)):
    da = combination[i][0]
    zipcode = combination[i][1] 
    distance, Zip_lat_long, b = compute_distance2(da, zipcode, Zip_lat_long)
    combination[i].append(distance)
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

This part is the first Optimization model (only last mile) that will remove DA
that are useless to achieve faster computation time for the last mile and line haul optimization

###############################################################
###############################################################
"""

#Create arcs out of combination if distance is less than threshold 
#   { Zip : {DA+Carrier :{distance, lm_cost (come in next step), var(come in two steps)}}
print("Build Arcs")

Arcs={}
for i in tqdm(range(nb_distances)):
    da = combination[i][0]
    zipcode = combination[i][1] 
    distance = combination[i][2]
#        Create an arc only if distance between DA and Zip is less than the Zip's state threshold
    if distance< Range[ZipCode_Dict[zipcode]['State']]:
        try:
            for carrier in DA_ZipCode_Dict[da]['Carrier']:
                Arcs[zipcode][da +" "+carrier ]={'distance' : distance}
        except KeyError:
            Arcs[zipcode]= {}
            for carrier in DA_ZipCode_Dict[da]['Carrier']:
                 Arcs[zipcode][da + " "+carrier]={'distance' : distance}

# Compute Costs for the arcs
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        
        distance = Arcs[pc][da]['distance']        
        da_state = DA_ZipCode_Dict[da[:5]]['State']
        da_carrier = da[6:]
        
        try:
            flat = Pricing[da_state][da_carrier]['Flat']
            breakpoint = Pricing[da_state][da_carrier]['Break']
            extra = Pricing[da_state][da_carrier]['Extra']
        except KeyError:
            sys.exit(("LM_Cost spreadsheet does not contain pricing info for couple state-carrier %s, %s" %(da_state,da_carrier)))
            
#        Check if distance Da_Zip is within flat distance
        if distance <  breakpoint:
            Arcs[pc][da]['lm_cost']=flat
        else :
            Arcs[pc][da]['lm_cost']=flat+ (distance - breakpoint) * extra
            
# Create Model
prob = pulp.LpProblem("Minimize Distance",pulp.LpMinimize)

# Design arcs
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        var = pulp.LpVariable("Arc_%s_%s)" % (pc,da),0,1,pulp.LpContinuous)
        Arcs[pc][da]['variable']=var

 

# Create Objective function : minimize distance
print("Create objective and Constraint")

#           We add a fraction of distance to the lm cost so we can avoid equality in price
prob += pulp.lpSum([(Arcs[pc][da]['lm_cost']+0.001*Arcs[pc][da]['distance'])*Arcs[pc][da]['variable'] for pc in Arcs.keys() for da in Arcs[pc].keys()])

# Create Constraint : every Zip is allocated
print("Create contraint 'every zipcode is assigned to a DA'")
for pc in tqdm(Arcs.keys()):          
    prob += pulp.lpSum([Arcs[pc][da]['variable'] for da in Arcs[pc].keys()]) == 1

# The problem is solved using PuLP's choice of Solver
print("Solve Problem")
start_time = time.clock()
prob.solve()
end_time = time.clock()
print(end_time-start_time)

# The status of the solution is printed to the screen
print("Status:", pulp.LpStatus[prob.status])


# The optimised objective function value is printed to the screen    
print ("total cost", pulp.value(prob.objective))

#Create workbook for results
w_result = xl.Workbook()
wresult = w_result.create_sheet('Optimization Results')
# export results on excel

print("Exporting Results")

wresult.cell(row=1,column=1).value= "ZipCode"
wresult.cell(row=1,column=2).value= "Carrier"
wresult.cell(row=1,column=3).value= "DaZipCode"
wresult.cell(row=1,column=4).value= 'DA and Carrier'
wresult.cell(row=1,column=5).value= 'Volume'
wresult.cell(row=1,column=6).value= 'Unit Cost'
# Print Results on excel
r=2
for pc in Arcs.keys():
    for da in Arcs[pc].keys():
        if Arcs[pc][da]['variable'].varValue !=0:
            wresult.cell(row=r,column=1).value= pc
            wresult.cell(row=r,column=2).value= da[6:]
            wresult.cell(row=r,column=3).value= da[:5]
            wresult.cell(row=r,column=4).value= da
            wresult.cell(row=r,column=5).value= ZipCode_Dict[pc]['Volume']            
            wresult.cell(row=r,column=6).value= Arcs[pc][da]['lm_cost']
            r+=1


print("Save File")

w_result.save("C:\HomeDepot_Excel_Files\Optimized.xlsx")


