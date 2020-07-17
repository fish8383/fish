import pandas as pd
a = input('分析表格名称？？')
df = pd.read_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+a+'.xlsx')

Tdata = pd.DataFrame(pd.to_datetime(df['2593IS0602-To WAX']) - pd.to_datetime(df['2593IS0208-FIN2.2 inlet'])).astype('timedelta64[s]')
T2data = pd.DataFrame(pd.to_datetime(df['2493IS0602-To WAX']) - pd.to_datetime(df['2493IS0206-FIN1.1 inlet'])).astype('timedelta64[s]')
# T3data = pd.DataFrame(pd.to_datetime(df['16IS200']) - pd.to_datetime(df['13IS065'])).astype('timedelta64[h]')
L1= df.copy()
L2=df.copy()
L2['F2']=Tdata

L1['F1']=T2data
print(L1)
# df['F2']=df['F2'].map(lambda x: x=0 if x>=4000000 else x )nb
L2['F2'] = L2['F2'].map(lambda x: 0 if x> 4000000 or x<0 else x)
L1['F1'] = L1['F1'].map(lambda x: 0 if x> 4000000  or x <0 else x)
L1=L1[~L1['F1'].isin([0])]
L2=L2[~L2['F2'].isin([0])]

# df['打磨到报交']=T3data
df2=L2.groupby(['Body type']).agg(['min','max','mean'])
df4=L1.groupby(['Body type']).agg(['min','max','mean'])
df.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'清单.xlsx')
df2.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'shop2.xlsx')
df4.to_excel('E:\DATA_ENGIN'+'\\'+'bodydata'+'\\'+'shop1.xlsx')