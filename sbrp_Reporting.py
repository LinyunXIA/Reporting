import pandas as pd  #vlookup
import openpyxl
import glob
import os

 #=======合并
path = "D:\\python\\files\\tmp\\"
new_workbook = Workbook()
new_sheet = new_workbook.active
 
 
# 用flag变量明确新表是否已经添加了表头，只要添加过一次就无须重复再添加 
# 0 or 1
flag = 1
 
 
for file in glob.glob(path + '/*.xlsx'):
    workbook = load_workbook(file)
    sheet = workbook.active
 
 
    coloum_A = sheet['A']
    row_lst = []
    for cell in coloum_A:
        if cell:
            print(cell.row)
            row_lst.append(cell.row)
 
 
    if not flag:
        header = sheet[1]
        header_lst = []
        for cell in header:
            header_lst.append(cell.value)
        new_sheet.append(header_lst)
        flag = 1
 
 
    for row in row_lst:
        data_lst = []
        for cell in sheet[row]:
            data_lst.append(cell.value)
        new_sheet.append(data_lst)
 
 
new_workbook.save(path + '/' + 'temp.xlsx')

#---筛选比对 vlookup

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

