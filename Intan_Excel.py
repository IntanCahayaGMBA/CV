#
# James, 2025/10/08
# File: James_Excel.py
# Calculate sum of column in Excel file.


import pandas as pd
import os

# 1. Input
df = pd.read_excel ('Output.xlsx')

# 2. process
sums = df.select_dtypes(include='number').sum()
df_with_sum = pd.concat([df, pd.DataFrame([sums])], ignore_index=True)

folder = r'C:\Users\62896\BAAI\BAAI'
files = sorted([f for f in os.listdir(folder) if f.endswith('.xlsx')])

CurrentFile = 'Output.xlsx'

if CurrentFile in files:
    print(f"\nReading: {CurrentFile}")
    df = pd.read_excel(os.path.join(folder, CurrentFile))
    print(df.head())

    if files.index(CurrentFile) == len(files) - 1:
        print("This is the last file.")
    else:
        print(f"Next file will be: {files[files.index(CurrentFile)+1]}")
