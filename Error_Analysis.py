# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 14:10:42 2017

@author: Steven Gao
"""

#Use Pandas to import excel file
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#Extract Data from Excel
ltl_price = pd.read_excel('C:\HomeDepot_Excel_Files\ltl_price.xlsx')
model_data = pd.read_excel('C:\HomeDepot_Excel_Files\Model_Output.xlsx')

#Create lists to extract excel file
invoice_id = []
apro_cost = []
orig_nm = []
tot_shp_wt = []
tot_mile_cnt = []

variables_1 = []
coeff_1 = []
variables_1 = []
coeff_1 = []
variables_1 = []
coeff_1 = []
variables_1 = []
coeff_1 = []
variables_1 = []
coeff_1 = []

predict_values = [];
temp_variable = [];
percent_error = [];
percent_error_filter = []
mean_error = [];

invoice_id_Data = ltl_price.iloc[0:len(ltl_price), 0]
apro_cost_Data = ltl_price.iloc[0:len(ltl_price), 7]
orig_nm_Data = ltl_price.iloc[0:len(ltl_price), 17]
tot_shp_wt_Data = ltl_price.iloc[0:len(ltl_price), 13]
tot_mile_cnt_Data = ltl_price.iloc[0:len(ltl_price), 14]

variables_Data = model_data.iloc[0:len(model_data), 0]
coeff_Data = model_data.iloc[0:len(model_data), 1]

#Change Data Structure
for i in range(0, len(invoice_id_Data)):
    invoice_id.append(invoice_id_Data.get_value(i))
    apro_cost.append(apro_cost_Data.get_value(i))
    orig_nm.append(orig_nm_Data.get_value(i))
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
        if (orig_nm[x] == variables[y]):
            temp_variable = coeff[y]
        else: 
            temp_variable = 0
    predict_values.append(coeff[0] + temp_variable + tot_shp_wt[x]*0.137 +tot_mile_cnt[x]*0.046 + tot_shp_wt[x]*tot_mile_cnt[x]*0.000085)

#Calculate Percent Error
for z in range(0, len(predict_values)):
    percent_error.append(100*(predict_values[z] - apro_cost[z])/predict_values[z])


for w in range(0, len(percent_error)):
    if (percent_error[w] >= -80):
        percent_error_filter.append(percent_error[w])

#Print Result
mean_error = np.mean(percent_error)
median_error = np.median(percent_error)
min_error = np.min(percent_error)
max_error = np.max(percent_error)
std_error = np.std(percent_error)

print ("mean error: ", mean_error)
print("median error: ",median_error)
print("min error: " , min_error)
print("max error: " , max_error)
print("std error: " , std_error)

#Plot
plt.scatter(range(0,len(percent_error)),percent_error)
plt.hist(percent_error_filter)
plt.xlim(-80,80)
plt.ylabel('Percent Error')
plt.show()


