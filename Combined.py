# -*- coding: utf-8 -*-
"""
Created on Fri Sep 29 14:53:36 2017
This file combine all other Python files with some improvments to avoid useless processes
This is an optimization of the Home Depot .com Delivery network, with objective function being the total cost (line haul + last Mile) 
@author: Cyprien Bastide, Steven (Gao) Ming, Edson David Silva Moreno
"""



#Import the differents Modules that are going to be used in this file
import openpyxl as xl
from Procedures import neig_states, cell, instance, compute_distance, correct_zip
from progress.bar import IncrementalBar


# Open Worksheet that contains list of DA, List of Zip code with Volume, Pricing spreadsheet,...
wb = xl.load_workbook('C:\HomeDepot_Excel_Files\Standard_File.xlsx')




# Open All Different Spreadsheet
w_neig = wb['List_of_Neighboring_States']
w_da = wb['DA_List']
w_zip = wb['Zip_Allocation_and_Pricing']
#wb.create_sheet('Distances')
#w_dis = wb['Distances']
#w_dis.cell(row=1,column=1).value = 'DA'
#w_dis.cell(row=1,column=2).value = 'ZipCode'
#w_dis.cell(row=1,column=3).value = 'Distance DA-Zip'




# Get number of DA and of Zip
n_da= instance(w_da)
n_zip = instance(w_zip)




# Create dictionnaries for DA and Zip(with volume in tuple) Grouped by State, 
# This is useful because Arcs are going to be created based on neighgboring states
# Since distances don't depend on Carrier, if multiple DA are in the same zipcode only one will be counted
State_Da_dict = {}
State_Zip_dict = {}

# Dictionnary for DA  {State : [Da_ZipCode]} 
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
for r in range(n_zip):
    state = cell(w_zip,r+2,2)
    zipcode = correct_zip(str(cell(w_zip,r+2,1))) 
    volume = cell(w_zip,r+2,7)
    try:
        State_Zip_dict[state].append((zipcode, volume))
    except KeyError:
        State_Zip_dict[state]=[]
        State_Zip_dict[state].append((zipcode, volume))
        
        
        
        
        
# Create couple that need to have distance assigned to it based Neighboring states, 
# Go through every state, look at the Da inside, and assign them to all zip code in neighboring states
line1=2
line2=2
combination = []
# Iterate through the states
for state in State_Da_dict.keys():
    print('Creating combination for state %s' % state)
    
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
                

# Save file
print ("save file")
wb.save('C:\HomeDepot_Excel_Files\File_modified.xlsx')

#Compute distances
print("Compute Distances")
compute_distance('C:\HomeDepot_Excel_Files\File_modified.xlsx','Distances',1,2,3)
