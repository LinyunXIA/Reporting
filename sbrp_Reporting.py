import pandas as pd  #vlookup
import openpyxl
import glob
import os

#opera_result
dfoperaorg1=pd.read_excel('D:/python/files/opera.xlsx',sheet_name='Sheet1')

dfRMH=pd.read_excel('D:/python/files/RMH.xlsx',sheet_name='RPA SBRP')

dfoperamail=pd.read_excel('D:/python/files/RMH.xlsx',sheet_name='Details')

dataopera1=dfoperaorg1.merge(dfRMH, on=['Inncode'])

dataopera=dataopera1.merge(dfoperamail,on=['Inncode'])

#dataopera1.to_csv('D:/python/files/opera_result.csv')

dataopera.to_csv('D:/python/files/opera_result.csv')

#oasis_result
dfoasisorg1=pd.read_excel('D:/python/files/oasis.xlsx',sheet_name='Sheet1')

dfoasismail=pd.read_excel('D:/python/files/RMH.xlsx',sheet_name='Details')

dataoasis1=dfoasisorg1.merge(dfRMH, on=['Inncode'])

dataoasis=dataoasis1.merge(dfoasismail,on=['Inncode'])

dataoasis.to_csv('D:/python/files/oasis_result.csv')

