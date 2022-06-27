# -*- coding: utf-8 -*-
"""
Created on Tue Jun 21 13:21:21 2022

@author: abockelman
"""

import pandas as pd
import requests
from lxml import html
from zipfile import ZipFile
import zipfile as zp
import os
import os.path
import datetime  

days_ago = 0

x = requests.get('https://www.ercot.com/gridinfo/resource')
webpage = html.fromstring(x.content)
links = webpage.xpath('//a/@href')
link_ruc_xlsx =links[219]

ruc_xlsx= pd.ExcelFile(link_ruc_xlsx)
ruc_fnames = [sheet for sheet in ruc_xlsx.sheet_names if sheet.endswith('Capacities')]
for ruc_fnames in ruc_fnames:
    cap_xlsx = pd.read_excel(ruc_xlsx, sheet_name = ruc_fnames)
    
cap_xlsx.columns=cap_xlsx.iloc[0]

cap_xlsx = cap_xlsx[1:]


cap_xlsx =cap_xlsx.filter(regex='(?i)^(?!NaN).+', axis=1)
cap_xlsx  = cap_xlsx.drop(['GENERATION INTERCONNECTION PROJECT CODE'],axis=1)
cap_xlsx = cap_xlsx.dropna()
cap_xlsx = cap_xlsx.sort_index().reset_index(drop=True)  
cap_xlsx.index +=1


your_path =os.getcwd()
files = os.listdir(your_path)

for file in files:
   # specifying the zip file name
    file_name = file
  
    if zp.is_zipfile(file):
        # opening the zip file in READ mode
        with ZipFile(file_name, 'r') as zip:
        # printing all the contents of the zip file
            zip.printdir()
              
            # extracting all the files
            print('Extracting all the files now...')
            zip.extractall()
            print('Done!')
#create df with HE 1-24

hourending = list(range(1,25))
hourending =[' '+ str(y) +':00' for y in hourending]
hr_ruc = pd.DataFrame(columns = hourending) 

df = pd.DataFrame()
for file in files:
    file_name = file
    if (datetime.date.today() - datetime.timedelta(days=days_ago)).strftime("%Y%m%d") in file_name and ".csv" in file_name: #check for current date
        df_temp = pd.read_csv(file_name)
        df = df.append(df_temp)
        df = df.reset_index(drop=True)
        for i in range(len(df)):
            df['ResourceName'][i] = df['ResourceName'][i].strip()
            
unit_stupid_list = list()
for i in range(len(df)):
    if df['ResourceName'][i] in unit_stupid_list:
        pass
    elif df['ResourceName'][i] in cap_xlsx['UNIT CODE'].values:
        unit_stupid_list.append(df['ResourceName'][i])            
            
hr_ruc.insert(0, 'UNIT CODE', unit_stupid_list)     
hr_ruc = hr_ruc.merge(cap_xlsx, how ='left', on ='UNIT CODE')
for i in range(len(hr_ruc)):
    for j in range(1,25):
        for k in range(len(df)):
            if ' ' + str(j) + ":00" == df['HourEnding'][k] and hr_ruc['UNIT CODE'][i] == df['ResourceName'][k]:
                hr_ruc[hr_ruc.columns[j]][i] = hr_ruc[hr_ruc.columns[-1]][i] 

# y = hr_ruc.columns.tolist()
# cols = y[:1] + y[-2:] + y[1:-2]
# hr_ruc = hr_ruc[cols]

hr_ruc.iloc[:,range(1,25)] = hr_ruc.iloc[:,range(1,25)].fillna(0).astype(int)
hr_ruc.loc['Total'] = hr_ruc.iloc[:,range(1,25)].astype(int).sum()




output = 'Daily EROCT RUC' + str(datetime.date.today() - datetime.timedelta(days=days_ago)) +'.xlsx'
# if path.os.path.isfile(your_path+ '/'+output):
#     output.close()

with pd.ExcelWriter(output) as writer:
    cap_xlsx.to_excel(writer, sheet_name = 'Seasonal Capacities')    
    hr_ruc.to_excel(writer, sheet_name = 'Daily RUC')      


    
 

# for file in files:   
#       # specifying the zip file name
#       file_name = file
#       if datetime.date.today().strftime("%Y%m%d") not in file_name and zp.is_zipfile(file_name) and ".xlsx" not in file_name:
#                os.remove(file_name)
#  #     if ".csv"  in file_name :
#   #             os.remove(file_name)
                  
          
    
