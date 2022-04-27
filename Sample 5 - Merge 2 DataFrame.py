import pandas as pd
import os
import sys
file_a = input('Input file A path: ')
# Raise error if file does not exist
if not os.path.isfile(file_a):
    raise FileNotFoundError('File does not exist')

file_b = input('Input file B path: ')

# Raise error if file does not exist
if not os.path.isfile(file_b):
    raise FileNotFoundError('File does not exist')

key_a = input('Input key column for dataframe A: ')
if key_a == '':
    key_a = 'KeyA'
key_b = input('Input key column for dataframe B: ')
if key_b == '':
    key_b = 'KeyB'
#pd.read_excel(file_a, engine='openpyxl', sheet_name= Sheet_Name, skiprows = Key_Row-1, usecols = List_Col)
# Raise error if cannot read file
try:
    df_a = pd.read_excel(file_a, engine='openpyxl', sheet_name= 'Data')
except:
    raise FileNotFoundError('Cannot read file')

try:
    df_b = pd.read_excel(file_b, engine='openpyxl', sheet_name= 'Data')
except:
    raise FileNotFoundError('Cannot read file')

# Raise error if cannot merge
try:
    df_merged = pd.merge(df_a, df_b, left_on=key_a, right_on=key_b, how="outer")
    df_merged.dropna(how='all', axis=1, inplace=True)
except:
    raise FileNotFoundError('Cannot merge')

df_merged.to_excel("output.xlsx", engine='openpyxl', sheet_name='Data')  
print('Done')