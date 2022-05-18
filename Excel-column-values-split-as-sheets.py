import pandas as pd
import os
os.chdir('M:\Track Splits')
from pandas import ExcelWriter

filename = 'Booked appointments May 2022'
df = pd.read_excel(r"P:\Analysis_HI\Data Requests\2022\TrakCare Down Time May 2022\Booked appointments May 2022.xlsx")

cols = list(df.columns)

col_of_interest = 'Doctor'
col_of_interest_values = list(df[col_of_interest].unique())

# Filter within col_of_interest
search = ['Dr Ray Armstrong', 'Dr Christopher Holroyd'] 
df = df[df[col_of_interest].isin(search)]

sheet_names = list(df[col_of_interest].unique())

# To Excel object loop through values, filter df, write to sheet and name sheet as per filter value
w = ExcelWriter(f'{filename}_splits.xlsx')
for n in sheet_names:
    res = df[df['Doctor'] == n]
    res.to_excel(w, sheet_name=n, index=False)
w .save()   
