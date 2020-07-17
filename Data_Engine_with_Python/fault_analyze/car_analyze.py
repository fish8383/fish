# 对汽车投诉信息进行分析hw
import pandas as pd

result = pd.read_excel('hw.xlsx')
print(result)
# 将genres进行one-hot编码（离散特征有多少取值，就用多少维来表示这个特征）
result = result.drop('problem', 1).join(result.problem.str.get_dummies(','))
#result.to_excel('temp1.xlsx', index=False)
#result.problem.str.get_dummies(',').to_excel('temp2.xlsx', index=False)
#print(result.columns)
tags = result.columns[7:]
print(tags)
#result2 = result.problem.str.get_dummies(',')
#print(result2)
df = result
df1 = result.groupby(['brand'])['id'].agg(['count'])
print(df1.sort_values('count',ascending=False))
#print(result2)
#print(result2)
df2= result.groupby(['car_model'])['id'].agg(['count'])

print(df2.sort_values('count',ascending=False))
df3 = result.groupby(['brand','car_model'])['id'].agg(['count'])
print(df3)
# 通过reset_index将DataFrameGroupBy => DataFrame
#df2.reset_index(inplace=True)
#df2.to_csv('temp.csv')
#df2= df2.sort_values('count', ascending=False)
#print(df2)
#print(df2.columns)
#df2.to_csv('temp.csv', index=False)
#query = ('A11', 'sum')
#print(df2.sort_values(query, ascending=False))