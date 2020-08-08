import numpy as np 
import pandas as pd 
import matplotlib as plt  
import datetime
import time
import openpyxl

#创建一个新的Excel表格
wb=openpyxl.Workbook()
sheet=wb.active
sheet['A1']='Name'
sheet['B1']='Total Watch Time'
#05Excel
data=pd.read_excel('05.xlsx')
ds=pd.DataFrame(data)
#05aExcel
data1=pd.read_excel('05a.xlsx')
ds1=pd.DataFrame(data1)
#06Excel
data2=pd.read_excel('06.xlsx')
ds2=pd.DataFrame(data2)
#07Excel
data3=pd.read_excel('07.xlsx')
ds3=pd.DataFrame(data3)
#07aExcel
data4=pd.read_excel('07a.xlsx')
ds4=pd.DataFrame(data4)
#利用numpy找出所有姓名和时间，并组成一个新的numpy
#第一个
ds['Total Watch Time']=ds['Total Watch Time'].map(lambda x: str(x))
ds['Total Watch Time']=ds['Total Watch Time'].map(lambda x: x.replace('无','00:00:00'))
aar=np.array([ds['Name']])
#新建list并将姓名导入list内部，用以后续的导入excel表格(i是numpy表格内部的列表，表明numpy可以转化为列表并写入Excel)
list=[]
for i in aar:
    for l in i:
        list.append(l)
ffr=ds['Total Watch Time']
#第二
ds1['Total Watch Time']=ds1['Total Watch Time'].map(lambda x: str(x))
ds1['Total Watch Time']=ds1['Total Watch Time'].map(lambda x: x.replace('无','00:00:00'))
bbr=np.array([ds1['Name']])
ggr=ds1['Total Watch Time']
#第三个
ds2['Total Watch Time']=ds2['Total Watch Time'].map(lambda x: str(x))
ds2['Total Watch Time']=ds2['Total Watch Time'].map(lambda x: x.replace('无','00:00:00'))
ccr=np.array([ds2['Name']])
hhr=ds2['Total Watch Time']
#第四个
ds3['Total Watch Time']=ds3['Total Watch Time'].map(lambda x: str(x))
ds3['Total Watch Time']=ds3['Total Watch Time'].map(lambda x: x.replace('无','00:00:00'))
ddr=np.array([ds3['Name']])
iir=ds3['Total Watch Time']
#第五个
ds4['Total Watch Time']=ds4['Total Watch Time'].map(lambda x: str(x))
ds4['Total Watch Time']=ds4['Total Watch Time'].map(lambda x: x.replace('无','00:00:00'))
eer=np.array([ds4['Name']])
jjr=ds4['Total Watch Time']
#合并上述五个numpy.daaray表格形式（40个人五次上课的时间）
df=np.array([ffr,ggr,hhr,iir,jjr]).T
#将姓名的numpy列写入Excel中
#找出每个人的分均上课时长
list1=[]
for x in range(0,len(df)):
    sum=0
    for y in range(5):
        (h,m,s)=str(df[x][y]).split(':')
        t=3600*int(h)+60*int(m)+int(s)
        sum+=t 
    #sum=str(datetime.timedelta(seconds=0))
    a=str(datetime.timedelta(seconds=sum))
    sheet.append([i[x],a])#(i因为是外部元素，且这一层属于x范围，因此写入i[x])

wb.save('Total time in 5.xlsx')