# The purpose of this python file is to find a regressional model that best fit the line haul data and
# output statistically significant terms and coefficients to an excel file


from statsmodels.formula.api import ols
import pandas as pd
import numpy as np

#Extract Data from Excel
ltl_price = pd.read_excel('C:\HomeDepot_Excel_Files\ltl_price.xlsx')
ltl_price.head()
ltl_price = ltl_price[(ltl_price["tot_shp_wt"] >= 200) & (ltl_price["tot_shp_wt"] <= 4999) & (ltl_price["aprv_amt"] <=1000)]
#ltl_price["tot_mile_cnt1"] = ltl_price["tot_mile_cnt"] - np.mean(ltl_price["tot_mile_cnt"])
#ltl_price["tot_shp_wt1"] = ltl_price["tot_shp_wt"] - np.mean(ltl_price["tot_shp_wt"])
ltl_price['tot_mile_wt'] = ltl_price["tot_mile_cnt"] * ltl_price["tot_shp_wt"]


# fit our model with .fit() and show results
linehaul_model = ols("aprv_amt ~ tot_mile_cnt + tot_shp_wt + orig_state + tot_mile_wt", data=ltl_price).fit()
# summarize our model
linehaul_model_summary = linehaul_model.summary()
print(linehaul_model_summary)

#Terms & Coeff
variables = ['intercept','GA','MD','OH','tot_shp_wt','tot_mile_cnt','tot_mile_wt']
coeff = [linehaul_model.params.Intercept,linehaul_model.params["orig_state[T.GA]"],linehaul_model.params["orig_state[T.MD]"],linehaul_model.params["orig_state[T.OH]"],linehaul_model.params["tot_shp_wt"],linehaul_model.params["tot_mile_cnt"],linehaul_model.params["tot_mile_wt"]]

# Convert two lists to a dictionary
dictionary = dict(zip(variables,coeff))

# Output to an excel file
df = pd.DataFrame({'Variables': variables,'Coeff': coeff})
# Fix column names
df = df[['Variables','Coeff']]
writer = pd.ExcelWriter('C:\HomeDepot_Excel_Files\Model_Output.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()




