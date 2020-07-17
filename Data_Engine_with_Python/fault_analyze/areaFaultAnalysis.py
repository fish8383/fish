import pandas as pd
import os
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta
from pandas import Series, DataFrame
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdate
from pylab import plot, show
import numpy as np
import pygal

inp = [{'c1':10, 'c2':100}, {'c1':11,'c2':110}, {'c1':12,'c2':120}]

df = pd.DataFrame(inp)

print (df)



 for i in df.index
     df.loc[i]=df.loc[i]*2
 print(df)

#df['c2']=df['c2'].map(lambda x:x*6)

df.loc[0]=df.loc[0]*2

print(df.loc[0])
print(df)