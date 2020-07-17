
import pandas as pd
import os
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta
from pandas import Series, DataFrame
import matplotlib
import matplotlib.pyplot as plt
from pylab import plot, show
import numpy as np

result = pd.read_excel('total.xlsx')

df= result


df1 = result.groupby(['提交人'])['故障描述'].agg(['count'])
# print(type(df1))
#print(df1.sort_values('count',ascending=False))
df1 = df1.sort_values('count',ascending=False)
df1.to_excel('提交人数量排列.xlsx')
df2 = result.groupby(['部门'])['故障描述'].agg(['count'])
df_noc = result.drop(columns=['故障代码']) 

df_noc=df_noc[~df_noc['严重度'].isin(['C'])]

df_noc.to_excel('AB故障.xlsx')
df2=(df2.sort_values('count',ascending=False))

df2.to_excel('组织提交数量排列.xlsx')

#计算单条故障的累计时间
startime = {}
closetime = {}
dettime = {}
startime = df_noc.loc[:,['结束时间']]
closetime = df_noc.loc[:,['开始时间']]
#开始结束时间差
new_df = pd.DataFrame(pd.to_datetime(df_noc['结束时间']) - pd.to_datetime(df_noc['开始时间']))
#day = pd.DataFrame(pd.to_datetime(df3['提交时间'])
col_name = df_noc.columns.tolist() 
col_name.insert(6,'stoptimer') 
#col_name.insert(7,'fault_date') 
# 在列索引为2的位置插入一列,列名为:故障时间差，刚插入时不会有值，整列都是NaN
df_noc=df_noc.reindex(columns=col_name)
#将时间差放入故障清单表
df_noc['stoptimer']=new_df  

#******************按日期计数*************#
#改英文
df_noc.rename(columns={'提交时间':'fault_date'}, inplace = True)

df_noc.rename(columns={'status ':'status'}, inplace = True)
#更改格式为DATE TIME##
day = pd.DataFrame(pd.to_datetime(df_noc['fault_date']))
# 将表格中的至改掉
df_noc['fault_date'] = day
df_noc['fault_date']= df_noc.fault_date.dt.date
#print(df_noc)
#########********日故障计数*************************************####################
daycount = df_noc.fault_date.value_counts()

#print(daycount) 
df_DayCount=df_noc.groupby(['fault_date'])['故障描述'].agg(['count'])
###*********************计算时间################

df_DayDowntime=df_noc[['fault_date','stoptimer']]
df_DayDowntime=(df_DayDowntime.sort_values('fault_date',ascending=False))
#print(df_DayDowntime)
df_DayDowntime=df_DayDowntime.groupby(['fault_date'])['stoptimer'].agg(['sum'])
#df_DayDowntime.rename(columns={'sum':'stoptimer'},inplace=True)

#######增加区域故障###########

#df_areaDonwtime=df_noc[['fault_date','stoptimer','部门']].copy
df_areaDonwtime = df_noc.copy()

df_areaDonwtime.rename(columns={'部门':'aaa'}, inplace = True) 

df_areaDonwtime['stoptimer']=df_areaDonwtime['stoptimer'].apply(timedelta.total_seconds)

# #print(type(sec1))
#df_areaDonwtime['stoptimer']=sec1
df_areaDonwtime['stoptimer'] = df_areaDonwtime['stoptimer'].map(lambda x: x/60)
df_areaDonwtime=df_areaDonwtime[['fault_date','stoptimer','aaa']]
col_name = df_areaDonwtime.columns.tolist() 
col_name.insert(4,'PT') 
col_name.insert(5,'PVC') 
col_name.insert(6,'TC') 
col_name.insert(7,'WAX') 

df_areaDonwtime=df_areaDonwtime.reindex(columns=col_name)
df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-注蜡','WAX'] = df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-注蜡','stoptimer']
df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-PVC','PVC'] = df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-PVC','stoptimer']
df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-底面漆','TC'] = df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-底面漆','stoptimer']
df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-预处理&电泳','PT'] = df_areaDonwtime.loc[df_areaDonwtime.aaa =='维修工段-预处理&电泳','stoptimer']
df_areaDonwtime=df_areaDonwtime.drop(columns=['aaa'])
df_areaDonwtime=df_areaDonwtime.drop(columns=['stoptimer'])




#print(df_areaDonwtime)
df_areaDonwtimePVC=df_areaDonwtime.groupby(['fault_date'])['PVC'].agg(['sum'])
df_areaDonwtimeTC=df_areaDonwtime.groupby(['fault_date'])['TC'].agg(['sum'])
df_areaDonwtimePT=df_areaDonwtime.groupby(['fault_date'])['PT'].agg(['sum'])
df_areaDonwtimeWAX=df_areaDonwtime.groupby(['fault_date'])['WAX'].agg(['sum'])
#
# df_areaDonwtime1 = df_areaDonwtime1.rename(columns={'sum_x':'TC'}, inplace = True)
df_areaDonwtimePVC.rename(columns={'sum':'PVC'}, inplace = True)
df_areaDonwtimeWAX.rename(columns={'sum':'WAX'}, inplace = True)
df_areaDonwtimeTC.rename(columns={'sum':'TC'}, inplace = True)
df_areaDonwtimePT.rename(columns={'sum':'PT'}, inplace = True)
df_areaDonwtime1 = pd.merge(df_areaDonwtimeTC,df_areaDonwtimeWAX, on = 'fault_date',how='left')
df_areaDonwtime2 = pd.merge(df_areaDonwtimePT,df_areaDonwtimePVC, on = 'fault_date',how='left')
df_areaDonwtime =  pd.merge(df_areaDonwtime1,df_areaDonwtime2, on = 'fault_date',how='left')



###*********************计算时间################

##转换格式#####

#####合并表格###################

#print(df_MTD)
######将DELTATIME转换##########

df_MTD = pd.merge(df_DayCount, df_DayDowntime, on = 'fault_date',how='left')


sec = df_MTD['sum'].apply(timedelta.total_seconds)


df_MTD['sum']=sec
df_MTD['sum'] = df_MTD['sum'].map(lambda x: x/60)

############******求平均停线时间*******################
df_MTD['mean'] = df_MTD['sum'].div(df_MTD['count'])
df_MTD.rename(columns={'sum':'TotalTime(min)','mean':'MTBF(min)','count':'TotalCount'}, inplace = True)
#print(df_MTD)
df_DayDowntime2s= df_DayDowntime.drop(columns=['sum'])
DayDowntime2s = df_MTD['TotalTime(min)']
df_DayDowntime2s['TotalTime(min)'] =DayDowntime2s


##################*******总表的停线时间格式更改，方便查看*****##############
sec = df_noc['stoptimer'].apply(timedelta.total_seconds)
df_noc['stoptimer'] = sec
df_noc['stoptimer'] = df_noc['stoptimer'].map(lambda x: x/60)



###############********总表中已关闭故障********########################
# print(df_noc)

df_status = df_noc[['fault_date','status']]
col_name = df_status.columns.tolist() 
col_name.insert(4,'OPN')
col_name.insert(5,'CLS') 
col_name.insert(6,'PND') 
df_status=df_status.reindex(columns=col_name) 


df_status.loc[df_status.status  =='closed','CLS'] = df_status.loc[df_status.status  =='closed','status']
df_status.loc[df_status.status  =='open','OPN'] = df_status.loc[df_status.status  =='open','status']
df_status.loc[df_status.status  =='pending','PND'] = df_status.loc[df_status.status  =='pending','status']

df_status_Count1 = df_status.groupby(['fault_date'])['CLS'].agg(['count'])
df_status_Count1.rename(columns={'count':'CLS'},inplace=True)
df_status_Count2 = df_status.groupby(['fault_date'])['OPN'].agg(['count'])
df_status_Count2.rename(columns={'count':'OPN'},inplace=True)
df_status_Count3 = df_status.groupby(['fault_date'])['PND'].agg(['count'])
df_status_Count3.rename(columns={'count':'PND'},inplace=True)
df_status = pd.merge(df_status_Count1,df_status_Count2,on = 'fault_date',how='left')
df_status =pd.merge(df_status,df_status_Count3,on='fault_date',how='left')


#####输出表格****************************################
df_noc.to_excel('累计故障.xlsx')
df_DayCount.to_excel('日故障计数.xlsx')
df_DayDowntime.to_excel('日停线时间.xlsx')
df_MTD.to_excel('MTBF平均停线时间.xlsx')
df_areaDonwtime.to_excel('区域停线时间.xlsx')
####图表输出############################################
import pygal
def histogram(xvalues,yvalues,a,b,C):
    hist= pygal.Bar()
    
    hist._title =b
    hist._x_title =a
    hist._y_title='频率'
    figure = plt.figure(dpi=100,figsize=(1000,600))
    hist.x_labels = xvalues
    hist.add(b,yvalues)
  
    hist.x_label_rotation = C
    hist.render_to_file(str(b)+'bar.svg')
xvalues =list(df_MTD.index)
yvalues = df_MTD['TotalCount']


histogram(xvalues,yvalues,'故障','faultcount',90) 
xvalues =list(df_DayDowntime2s.index)
yvalues = df_DayDowntime2s['TotalTime(min)']
histogram(xvalues,yvalues,'故障','faultduring',90)

####线图#########
import matplotlib.dates as mdate
def plot_curve1(data,title):
    fig1 = plt.figure(figsize=(15,5))
    ax1 = fig1.add_subplot(1,1,1)
    ax1.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
    #设置时间标签显示格式
    plt.xticks(list(df_MTD.index),rotation=45, fontsize =15) 
    plt.title(title)

    plt.plot(data,'o-')
    plt.savefig(title+'curv.svg')   



def plot_curve4(data1,data2,data3,data4,title):
    fig1 = plt.figure(figsize=(15,5))
    ax1 = fig1.add_subplot(1,1,1)
    ax1.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
    #设置时间标签显示格式
    plt.xticks(list(df_MTD.index),rotation=90)
    plt.title(title)

    plt.plot(data1,'o-')
    plt.plot(data2,'o-')
    plt.plot(data3,'o-')
    plt.plot(data4,'o-')
    plt.legend()
    plt.savefig(title+'curve.svg')
#############################组合柱状图*******************#####################################    
  ###########################################################################################

x = np.array(df_status.index)
y1 = np.array(list(df_status['CLS']))
y2 = np.array(list(df_status['OPN']))
y3 = np.array(list(df_status['PND']))

# plt.bar(x, y1, label="close", color='green')
# plt.bar(x, y2, label="open",color='red')
# plt.bar(x, y3, label="pending", color='yellow')
# plt.bar(x, y1,  width=width, label='label1',color='red')
# plt.bar(x + width, y2, width=width, label='label2',color='deepskyblue')
# plt.bar(x + 2 * width, y3, width=width, label='label3', color='green')
plt.barh(x, y1, color='green', label='closed')
plt.barh(x, y2, left=y1, color='red', label='Open')
plt.barh(x, y3, left=y1+y2, color='blue', label='pending')
plt.xticks(rotation=90, fontsize=10)  # 数量多可以采用270度，数量少可以采用340度，得到更好的视图
plt.legend(loc="upper right")  # 防止label和图像重合显示不出来
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.xlabel('完成情况')
plt.ylabel('故障日期')
plt.rcParams['savefig.dpi'] = 100  # 图片像素
plt.rcParams['figure.dpi'] = 100  # 分辨率
plt.rcParams['figure.figsize'] = (155.0, 155.0)  # 尺寸
plt.title("故障措施完成情况")
show()
#plt.savefig('故障完成情况.png') 
#plt.savefig('D:\\result.png')



plot_curve4(df_areaDonwtime['PT'],df_areaDonwtime['TC'],df_areaDonwtime['PVC'],df_areaDonwtime['WAX'],'FaultDuring_Area')

plot_curve1(df_DayDowntime2s,'FaultDuring_Total')

plot_curve1(df_DayCount,'FaultCount_Total')
# df4.plot(linewidth=1.0,color = 'yellow')
# plt.title("total_downcounter")
# #line


# df_DayDowntime.plot(linewidth=1.0,color = 'red')
# plt.title("total_downtime")

# show()

print('完成故障数据处理')

