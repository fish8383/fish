import pandas as pd
a=input('表格位置')
df = pd.read_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+a+'\\shop.xlsx')
Tdata = pd.DataFrame(pd.to_datetime(df['K2IS075']) - pd.to_datetime(df['01IS045'])).astype('timedelta64[m]')
T2data = pd.DataFrame(pd.to_datetime(df['10IS335']) - pd.to_datetime(df['01IS045'])).astype('timedelta64[m]')
T3data = pd.DataFrame(pd.to_datetime(df['16IS305']) - pd.to_datetime(df['10IS335'])).astype('timedelta64[m]')
T4data = pd.DataFrame(pd.to_datetime(df['K2IS075']) - pd.to_datetime(df['16IS305'])).astype('timedelta64[m]')
T5data = pd.DataFrame(pd.to_datetime(df['K2IS075']) - pd.to_datetime(df['01IS205'])).astype('timedelta64[m]')
L1= df.copy()
L2=df.copy()
L3=df.copy()
L4=df.copy()
L5=df.copy()

L1['L8-L1']=Tdata
L2['进口到PVC烘房']=T2data
L3['面漆到PVC烘房']=T3data
L4['面漆到L8']=T4data
L5['L8-L100_2']=T5data
# print(df['L8-L1'])
# df.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'shop1.xlsx')


print(L1)
# df['F2']=df['F2'].map(lambda x: x=0 if x>=4000000 else x )nb
L1['L8-L1'] = L1['L8-L1'].map(lambda x: 0 if x> 4000000 or x<0 else x)
L2['进口到PVC烘房'] = L2['进口到PVC烘房'].map(lambda x: 0 if x> 4000000  or x <0 else x)
L3['面漆到PVC烘房'] =L3['面漆到PVC烘房'].map(lambda x: 0 if x> 4000000  or x <0 else x)
L4['面漆到L8'] =L4['面漆到L8'].map(lambda x: 0 if x> 4000000  or x <0 else x)
L5['L8-L100_2'] =L5['L8-L100_2'].map(lambda x: 0 if x> 4000000  or x <0 else x)
L1=L1[~L1['L8-L1'].isin([0])]
L2=L2[~L2['进口到PVC烘房'].isin([0])]
L3=L3[~L3['面漆到PVC烘房'].isin([0])]
L4=L4[~L4['面漆到L8'].isin([0])]
L5=L5[~L5['L8-L100_2'].isin([0])]

# df['打磨到报交']=T3data
df1=L1.groupby(['Body Type']).agg(['min','max','mean'])
df2=L2.groupby(['Body Type']).agg(['min','max','mean'])
df3=L3.groupby(['Body Type']).agg(['min','max','mean'])
df4=L4.groupby(['Body Type']).agg(['min','max','mean'])
df5=L5.groupby(['Body Type']).agg(['min','max','mean'])
""" df1= df1.drop(df1.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12]],axis=1)
df2= df2.drop(df1.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12]],axis=1)
df3= df3.drop(df1.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12]],axis=1)
df4= df4.drop(df1.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12]],axis=1)
df5= df5.drop(df1.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12]],axis=1) """
df.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'清单.xlsx')
df2.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'进口到PVC烘房.xlsx')
df4.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'面漆到L8.xlsx')
df3.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'面漆到PVC烘房.xlsx')
df1.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'L8-L1.xlsx')
df5.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'L8-L1——2.xlsx')