import os
import pandas as pd
import numpy as np

os.chdir('C:\\Users\\Ben\\Documents\\Programming\\PythonHello\\Attachments\\2021\\Q1\\EU\\March')
cwd = os.path.abspath('')
files = os.listdir(cwd)

data = pd.read_excel('total_calls.xlsx')
df = pd.DataFrame(data)
df = df.rename(columns={'[Incoming Calls]': 'IncCall','[Outgoing Calls (Manual)]': 'OutCall', '[Organisation Hierarchy]': 'Org', '[Incoming Emails]': 'IncEmail', '[SMS Messages]': 'SMS', '[Total Social Media Messages]': 'Social', '[Connected Web Chats]': 'Webchat', '[OUTBOUND Connected Calls]': 'Dial'})

df['IncCall'] = df.IncCall.str.split('(').str[-1].str.strip(')').astype(int)
df['OutCall'] = df.OutCall.str.split('(').str[-1].str.strip(')').astype(int)
df['IncEmail'] = df.IncEmail.str.split('(').str[-1].str.strip(')').astype(int)
df['SMS'] = df.SMS.str.split('(').str[-1].str.strip(')').astype(int)
df['Social'] = df.Social.str.split('(').str[-1].str.strip(')').astype(int)
df['Webchat'] = df.Webchat.str.split('(').str[-1].str.strip(')').astype(int)
df['Dial'] = df.Dial.str.split('(').str[-1].str.strip(')').astype(int)

df = df.astype({'IncCall': np.float64,'OutCall': np.float64, 'IncEmail': np.float64, 'SMS': np.float64, 'Social': np.float64,'Webchat': np.float64,'Dial': np.float64})

table = pd.pivot_table(df, values=['IncCall','OutCall','IncEmail','SMS','Social','Webchat','Dial'], index='Org', aggfunc=np.sum, fill_value=0)

table.to_excel('sumtotal_calls.xlsx')
