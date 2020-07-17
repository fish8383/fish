###################################################
""" 2020年7月9日
将中控室导出的过车序列表格（英文），生成过点时间差。生成文件名按过点位置命名





 """
 ###########################################################
import pandas as pd
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta
createVar = locals()
crdf= locals()
while True :
    a= input('需要打开的数量')
    a= int(a)
    if a==0:
        break
    for i in range(a):
        b= input('需要打开的文件')
    
        createVar['df'+str(i)]= pd.read_excel('E:\DATA_ENGIN\\bodydata\\file\\'+b+'.xlsx')
        C_name =createVar['df'+str(i)].iloc[1,7]
        col_name = createVar['df'+str(i)].columns.tolist() 
        col_name.insert(4,C_name) 
        createVar['df'+str(i)]=createVar['df'+str(i)].reindex(columns=col_name)
        createVar['df'+str(i)].iloc[:,4] = createVar['df'+str(i)].Timestamp
        
        createVar['df'+str(i)]= createVar['df'+str(i)].drop(createVar['df'+str(i)].columns[[0,1,11,5,6,7,8,9,10]],axis=1)
    break

df =createVar['df'+str(0)]
print(df)
# df=df.drop(df.columns[[1,5,6]],axis=1)
# df2 = pd.merge(df1, df,how='inner',on='BodyId')
# print(df)
for i in range(1,a):
    df=df.drop(['Body type'],axis=1)
    df= pd.merge(df,createVar['df'+str(i)] ,how='inner',on='BodyId')
    # crdf['delta'+str(i)]=df
    Area =df.copy()
    print(df)
    X= int(2+i) 
    y= int(i)
    deltaT = pd.DataFrame(pd.to_datetime(df.iloc[:,X]) - pd.to_datetime(df.iloc[:,y]))
    Area['DealtaA']=deltaT
    Area['DealtaA']=Area['DealtaA'].apply(timedelta.total_seconds).map(lambda x: x/60)
    Area['DealtaA']=Area['DealtaA'].map(lambda x : 0 if x>300 else x)
    Area=Area[~Area['DealtaA'].isin([0])]
    
    AreaOverview=Area.groupby(['Body type']).agg(['min','max','mean'])
    print(Area)
    #####取一下区间名字
    C_name =Area.columns.tolist()
    P_a = C_name[X]
    p_b = C_name[y]
    print(C_name)
    
    Area.to_excel('E:\DATA_ENGIN\\bodydata\\file\\'+P_a+p_b+'AreaDelta.xlsx')
    AreaOverview.to_excel('E:\DATA_ENGIN\\bodydata\\file\\'+P_a+p_b+'AreaOverview.xlsx')

df.to_excel('E:\DATA_ENGIN\\bodydata\\file\\'+b+'analysis.xlsx')