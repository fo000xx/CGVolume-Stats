import os
import pandas as pd

os.chdir('C:\\Users\\Ben\\Documents\\Programming\\PythonHello\\Attachments\\2021\\Q1\\EU\\March')
cwd = os.path.abspath('')
files = os.listdir(cwd)

df = pd.DataFrame()

for file in files:
    if file.endswith('.csv'):
       df = df.append(pd.read_csv(file), ignore_index=True)

df.to_excel('total_calls.xlsx')
        