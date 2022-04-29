from venv import create
from openpyxl import load_workbook, Workbook #合并
import glob #合并
import pandas as pd  # 比对
import datetime
import os
import shutil # file operation

 
#---变量
path_tmp = "D:\\python\\files\\tmp\\"
pathRMH = "D:\\python\\files\\"
path_folder="D:\\python\\files\\result\\"
day_folder=datetime.datetime.now().strftime('%Y')+"\\"+datetime.datetime.now().strftime('%m')+ "\\" + datetime.datetime.now().strftime('%d')
time_folder = datetime.datetime.now().strftime("%H")
new_workbook = Workbook()
new_sheet = new_workbook.active
mkdir_target_path = "D:\\python\\files\\result\\" + day_folder
new_target_path = mkdir_target_path +"\\"+ time_folder

# -- 新建文件夹
#if os.path.isdir(mkdir_target_path):
#    os.mkdir(os.path.join(mkdir_target_path,time_folder))
if not os.path.exists(new_target_path):
    os.makedirs(new_target_path)
 
# 用flag变量明确新表是否已经添加了表头，只要添加过一次就无须重复再添加 
# 0 or 1
flag = 1
 
 
for file in glob.glob(path_tmp + '/*.xlsx'):
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
  
new_workbook.save(path_tmp + '/' + 'temp.xlsx')

#---筛选比对 vlookup

dforg=pd.read_excel(path_tmp + '/' + 'temp.xlsx',sheet_name='Sheet')
dfRMH=pd.read_excel(pathRMH + '/' + 'RMH.xlsx',sheet_name='RPA SBRP')
dfmailbox=pd.read_excel(pathRMH + '/' + 'RMH.xlsx',sheet_name='Details')
data1=dforg.merge(dfRMH, on=['Inncode'])
data2=data1.merge(dfmailbox,on=['Inncode'])
data2.to_csv(path_tmp + '/' + 'result.csv',index=False)

#---移动文件
def move_file(old_path,new_path):
    print(old_path)
    print(new_path)
    filelist = os.listdir(old_path)
    print(filelist)
    for file in filelist:
        src = os.path.join(old_path, file)
        dst = os.path.join(new_path, file)
        print('src:',src)
        print('dst:',dst)
        shutil.move(src,dst)
 
if __name__=='__main__':
    move_file(path_tmp,new_target_path)