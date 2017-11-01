# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 14:10:42 2017

@author: Steven Gao
"""
#The purpose of this python file is to compare between the actual cost and the cost from regression model


#Use Pandas to import excel file
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#Extract Data from Excel
ltl_price = pd.read_excel('C:\HomeDepot_Excel_Files\Standard_File.xlsx', sheetname='ltl_price')
model_data = pd.read_excel('C:\HomeDepot_Excel_Files\Model_Output.xlsx')

#Create lists to extract excel file
invoice_id = []
apro_cost = []
orig_state = []
tot_shp_wt = []
tot_mile_cnt = []

variables = []
coeff = []

predict_values = [];
temp_variable = [];
percent_error = [];
percent_error_filter = []
mean_error = [];

invoice_id_Data = ltl_price.iloc[0:len(ltl_price), 0]
apro_cost_Data = ltl_price.iloc[0:len(ltl_price), 7]
orig_state_Data = ltl_price.iloc[0:len(ltl_price), 17]
tot_shp_wt_Data = ltl_price.iloc[0:len(ltl_price), 13]
tot_mile_cnt_Data = ltl_price.iloc[0:len(ltl_price), 14]

variables_Data = model_data.iloc[0:len(model_data), 0]
coeff_Data = model_data.iloc[0:len(model_data), 1]

#Change Data Structure
for i in range(0, len(invoice_id_Data)):
    invoice_id.append(invoice_id_Data.get_value(i))
    apro_cost.append(apro_cost_Data.get_value(i))
    orig_state.append(orig_state_Data.get_value(i))
    tot_shp_wt.append(tot_shp_wt_Data.get_value(i))
    tot_mile_cnt.append(tot_mile_cnt_Data.get_value(i))

#Change Data Structure
for i in range(0, len(variables_Data)):
    variables.append(variables_Data.get_value(i))
    coeff.append(coeff_Data.get_value(i))

    
# Append predicted values using a regressional model
for x in range(0, len(invoice_id)):
    for y in range(0, len(variables)):
        if (tot_shp_wt[x] < 200):
            tot_shp_wt[x] = 200
        else:
            tot_shp_wt[x] = tot_shp_wt[x]
        if (orig_state[x] == variables[y]):
            temp_variable = coeff[y]
        else: 
            temp_variable = 0
    predict_values.append(coeff[0] + temp_variable + tot_mile_cnt[x]*coeff[1] +tot_shp_wt[x]*coeff[2] + tot_shp_wt[x]*tot_mile_cnt[x]*coeff[3])

#Calculate Percent Error
for z in range(0, len(predict_values)):
    #if (tot_shp_wt[z] <= 4999):
        percent_error.append(100*(predict_values[z] - apro_cost[z])/predict_values[z])


#Print Result
mean_error = np.mean(percent_error)
median_error = np.median(percent_error)
min_error = np.min(percent_error)
max_error = np.max(percent_error)
std_error = np.std(percent_error)
actual_total_cost = np.sum(apro_cost)
predicted_total_cost = np.sum(predict_values)
total_error_percent = 100*(predicted_total_cost-actual_total_cost)/predicted_total_cost

print("Actual Total Cost: " , actual_total_cost)
print("Predicted Total Cost: " , predicted_total_cost)
print("Total Error Percentage: " , total_error_percent)
print ("mean error: ", mean_error)
print("median error: ",median_error)
print("min error: " , min_error)
print("max error: " , max_error)
print("std error: " , std_error)

#Plot
plt.figure()
#plt.scatter(range(0,len(percent_error)),percent_error)
#plt.hist(percent_error_filter)
plt.scatter(tot_shp_wt,apro_cost, color = 'k', label = 'Actual Cost')
plt.scatter(tot_shp_wt, predict_values,color = 'r', label = 'Predicted Cost')
plt.xlim(min(tot_shp_wt), max(tot_shp_wt)+100)
plt.legend(loc='upper right', fontsize=20)
#plt.ylabel('Percent Error')
plt.suptitle('Cost vs Weight', fontsize = 20)
plt.xlabel('Weight', fontsize = 20)
plt.ylabel('Cost', fontsize = 20)
plt.show()

