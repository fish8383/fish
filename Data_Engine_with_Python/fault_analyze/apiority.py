import pandas as pd
import time
pd.get_option('display.width')
pd.set_option('display.width',180)
# 数据加载
F=input('文件名')
data = pd.read_excel('E:\DATA_ENGIN\\apiority\\'+F+'.xlsx')
print(data)
# 统一小写
data['Alarm Msg.'] = data['Alarm Msg.'].str.lower()
data['Alarm Msg.'] = data['Alarm Msg.'].str.replace(r'[^\w\s]+', '')
data['Alarm Msg.'] = data['Alarm Msg.'].str.replace('<','less').astype('str')
# 去掉none项
data = data.drop(data[data['Alarm Msg.'] == 'none'].index)
# 提取PLC名和设备名
data['AreaL']=data['Area']
data['AreaR']=data['Area']
data['AreaL']=data['AreaL'].str.split('_',expand=True)
data['AreaR']=data['AreaR'].str.split('_',expand=True)[3]
a=1
# 添加故障发生序号

for i in range(len(data.iloc[:,1])-1):
	j=i+1
	
	if data.iloc[i,1]==data.iloc[j,1]:
		
		data.iloc[j,0]=a
		data.iloc[i,0]=a
		
		
	else:
		a=a+1
		data.iloc[j,0]=a
		
data.rename(columns={'Unnamed: 0':'FaultNo.'},inplace=True)
data.to_excel('E:\DATA_ENGIN\\apiority\c1.xlsx')
# 采用efficient_apriori工具包
def rule1():
	from efficient_apriori import apriori
	start = time.time()
	# 得到一维数组orders_series，并且将Transaction作为index, value为Item取值
	orders_series = data.set_index('FaultNo.')['Alarm Msg.']
	# 将数据集进行格式转换
	transactions = []
	temp_index = 0
	for i, v in orders_series.items():
		if i != temp_index:
			temp_set = set()
			temp_index = i
			temp_set.add(v)
			transactions.append(temp_set)
		else:
			temp_set = set()
			temp_set.add(v)
	
	# 挖掘频繁项集和频繁规则
	itemsets, rules = apriori(transactions, min_support=0.01,  min_confidence=0.2)
	itemsets = pd.DataFrame(itemsets)
	rules = pd.DataFrame(rules)
	print('频繁项集：', itemsets)
	print('关联规则：', rules)
	itemsets.to_excel('E:\DATA_ENGIN\\apiority\\'+F+'频繁项集.xlsx')
	rules.to_excel('E:\DATA_ENGIN\\apiority\\'+F+'关联规则.xlsx')   
	end = time.time()
	print("本次用时：", end-start)
    


def encode_units(x):
    if x <= 0:
        return 0
    if x >= 1:
        return 1
# 采用mlxtend.frequent_patterns工具包
def rule2():
	from mlxtend.frequent_patterns import apriori as ap
	from mlxtend.frequent_patterns import association_rules
	pd.options.display.max_columns=1000
	start = time.time()
	hot_encoded_df=data.groupby(['FaultNo.','Alarm Msg.'])['Alarm Msg.'].count().unstack().reset_index().fillna(0).set_index('FaultNo.')
	
	hot_encoded_df = hot_encoded_df.applymap(encode_units)
	frequent_itemsets = ap(hot_encoded_df, min_support=0.01, use_colnames=True)
    
	rules = association_rules(frequent_itemsets, metric="lift", min_threshold=0.2)
   
	print("频繁项集：", frequent_itemsets)
    
	print("关联规则：", rules[ (rules['lift'] >= 1) & (rules['confidence'] >= 0.2) ])
	print(rules['confidence'])
	rules.to_excel('E:\DATA_ENGIN\\apiority\\'+F+'关联规则2.xlsx')
	frequent_itemsets.to_excel('E:\DATA_ENGIN\\apiority\\'+F+'频繁项集2.xlsx')
	end = time.time()
	print("总用时：", end-start)
    
rule1()
print('-'*100)
rule2()
