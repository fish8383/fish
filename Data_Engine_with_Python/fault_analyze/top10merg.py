"  这个程序用于将每日的报警整理成每周，每月的报表        "

"   姜帆   2020年 7月10日  "




import pandas as pd
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta
createVar = locals()
crdf= locals()

while True :
    a= input('需要打开的数量?')
    a= int(a)
    if a==0:
        break
    for i in range(a):
        florder= input('文件夹')
        filename= input('文件名')
        createVar['df'+str(i)]= pd.read_excel('E:\DATA_ENGIN\\'+florder+'\\'+filename+'.xlsx')
    
    break
Save=input('保存位置')
df =createVar['df'+str(0)]
for i in range(1,a):

    df= pd.concat([df,createVar['df'+str(i)]])
df['Alarm Msg.']=df['Alarm Msg.'].map(lambda x: '' if type(x)==str and x[0]=='=' else x)   
df.to_excel('E:\DATA_ENGIN\\'+Save+'\\'+filename+'.xlsx')
deltaT = pd.DataFrame(pd.to_datetime(df['End']) - pd.to_datetime(df['Begin']))

df['Begin']= pd.DataFrame(pd.to_datetime(df['Begin']))
df['Duration']=deltaT
df=df[~df['Duration'].isin(['NAN'])]


df['Duration']=df['Duration'].apply(timedelta.total_seconds)
df['Duration']=df['Duration'].map(lambda x: x/60)

df['Fdate']= df.Begin.dt.date

df1= df.groupby(['Fdate'])['Begin'].agg(['count'])
df2= df.groupby(['Fdate'])['Duration'].agg(['sum','min','max','mean',])
df.to_excel('E:\DATA_ENGIN\\WeeklyOverview_'+filename+'.xlsx')
df1.to_excel('E:\DATA_ENGIN\\Frequence_'+filename+'.xlsx')
df2.to_excel('E:\DATA_ENGIN\\Duration'+filename+'.xlsx')