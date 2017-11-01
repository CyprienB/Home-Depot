# The purpose of this python file is to find a regressional model that best fit the line haul data and
# output statistically significant terms and coefficients to an excel file


from statsmodels.formula.api import ols
import pandas as pd

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
for i in range(0, len(variables)):
    if variables[i].find("orig_state") != -1:
        variables[i] = variables[i][-3]+variables[i][-2]
        
coeff = [linehaul_model.params.tolist()][0]
# Convert two lists to a dictionary
coefficient = dict(zip(variables,coeff))

# Output to an excel file
df = pd.DataFrame({'Variables': variables,'Coeff': coeff})
# Fix column names
df = df[['Variables','Coeff']]
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\Model_Output1.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()




