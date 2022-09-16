import pandas as pd
import os
from pandas import ExcelWriter

"""
A program that splits a DataFrame by a column's values and uses the said values to rename sheets
in Excel before writing the column splits.

"""
# define CWD
os.chdir('M:\Track Splits')

# path to file to split
path_to_file = "P:\Analysis_HI\Data Requests\2022\TrakCare Down Time May 2022\Booked appointments May 2022.xlsx"

# define file name for output file
file_name = os.path.basename(path_to_file).split(.)[0]

# file to split into sheets
df = pd.read_excel(r"P:\Analysis_HI\Data Requests\2022\TrakCare Down Time May 2022\Booked appointments May 2022.xlsx")

# columns to object
cols = list(df.columns)

# define column to split use as sheet name
col_of_interest = 'Doctor'

# if filtering of col_of_interest required, uncomment the following:
# search = ['Dr A', 'Dr B'] 
# df = df[df[col_of_interest].isin(search)]

# define sheet names
sheet_names = list(df[col_of_interest].unique())

# to Excel object loop through values, filter DataFrame, write to sheet and name sheet as per filter value
w = ExcelWriter(f'{file_name}_splits.xlsx')
for n in sheet_names:
    res = df[df['Doctor'] == n]
    res.to_excel(w, sheet_name=n, index=False)
w .save()
