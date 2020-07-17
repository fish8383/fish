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
import lxml
import tinycss
import cssselect
while True:
    """ 中控室故障导出后生成图表 2020年7月9日     """
    # import cairosvg
    ###############******数据清理*******#######################
    fileName =input('请输入要打开表格名称')
    if fileName is 'N':
        print('结束~')
        break
    saveAdress =input('请输入要打开的文件夹')
    df = pd.read_excel('E:\DATA_ENGIN'+'\\'+saveAdress+'\\'+fileName+'.xlsx')
    save='E:\DATA_ENGIN'+'\\'+saveAdress+'\\'
    #df = pd.read_excel('process.xlsx')

    """ df=df.drop(df.index[[0,1,2]],axis= 0)
    df.to_excel('E:\DATA_ENGIN'+'\\'+saveAdress+'\\'+fileName+'.xlsx')
    print(df) """
    ################******表格整理******#######################
    hours=df.copy()
    Adt = pd.DataFrame(pd.to_datetime(hours['Begin']))
    hours['Begin']= Adt
    hours['h']= hours['Begin'].dt.hour
    hours['day']= hours['Begin'].dt.date

    #################故障等级停线1 不停线2######################
    hours['Priority']=hours['Priority'].apply(lambda x:1 if int(x)<60 else 2)
    P60 = hours.copy()

    #################聚合每个小时，按小时拉曲线
    
    P60=P60[~P60['Priority'].isin(['1'])]
    
    P60=P60.groupby(['h'])['Begin'].agg(['count'])
    print(P60)
    hours=hours.groupby(['h'])['Begin'].agg(['count'])
    print (hours)
    ###1.新增一列PLC名####
    # df['AreL']=''
    # df['AreR']=''
    df['AreaL']=df['Area']
    df['AreaR']=df['Area']
    df['AreaL']=df['AreaL'].str.split('_',expand=True)
    df['AreaR']=df['AreaR'].str.split('_',expand=True)[3]




    ####2.

    ################排列方法1，时间   2，频次###################

    #####1，频次#####
    ###############*******1.按PLC区域排列***###################
    def str_2(sx):
        sx= sx[-2]
    df_PLC=df.groupby(['AreaL'])['Begin'].agg(['count'])
    df_PLC=df_PLC.sort_values('count',ascending=False)
    df_PLC =df_PLC.head(25)
    ###############*******2.按设备类型排列***###################
    import re
    df_type= df.copy()

    an = re.search("conveyor", fileName)
    if an:
        a= Series()
    
        a=df_type['AreaR'].str.extract(r'(([A-Z]){2})')
        df_type['AreaR']=a
  
    
    #####下面别动 
    df_type=df_type.groupby(['AreaR'])['Begin'].agg(['count'])
    df_type=df_type.sort_values('count',ascending=False)
    df_type = df_type.head(25)

    print(df_type)

    ###############*******3.按设备标签号排列***###################
    df_device=df.groupby(['Area'])['Begin'].agg(['count'])
    df_device=df_device.sort_values('count',ascending=False)
    df_device =df_device.head(25)
    ###############*******4.按故障文本排列***###################


    df_message=df.groupby(['Alarm Msg.'])['Begin'].agg(['count'])
    df_message=df_message.sort_values('count',ascending=False)
    df_message=df_message.head(25)

    ########################5.按日期次数排列##########################
    date= pd.to_datetime(df['Begin'])

    df_date =df.copy()
    date= date.dt.date
    df_date['Begindate']=date


    df_date_frq=df_date.groupby(['Begindate'])['Begin'].agg(['count'])
    df_date_frq=df_date_frq.sort_values('Begindate',ascending=False)

    print(df_date_frq)

    ##########2.时间###########
    def str2sec(x):
        '''
        字符串时分秒转换成秒
        '''
        h, m, s = x.strip().split(':') #.split()函数将其通过':'分隔开，.strip()函数用来除去空格
        return int(h)*3600 + int(m)*60 + int(s) #int()函数转换成整数运算

    #######################1.按PLC排列#########################
    df['Duration'] = df['Duration'].apply(str2sec)
    df['Duration'] = df['Duration'].map(lambda x:x/3600)
    df['Duration'] = df['Duration'].map(lambda x:float(x))
    df_PLCT=df.groupby(['AreaL'])['Duration'].agg(['sum'])

    df_PLCT=df_PLCT.sort_values('sum',ascending=False)
    df_PLCT =df_PLCT.head(30)
    ########################2.按日期时间长短排列#######################
    df_date_during = df.copy()
    date= pd.to_datetime(df_date_during['Begin'])


    date= date.dt.date
    df_date_during['Begin']=date

    df_date_during=df_date_during.groupby(['Begin'])['Duration'].agg(['sum'])



    ##########################3.TOP10 频次###################################

    df_top10 = df.copy()
    print(df_top10)
    df_top10['Area']=df['AreaL']+'/'+df['AreaR']+'/'+df['Alarm Msg.']
    df_top10=df_top10.groupby(['Area'])['Begin'].agg(['count'])
    df_top10=df_top10.sort_values('count',ascending=False)
    df_top10 =df_top10.head(20)

    df_top10During =df.copy()

    df_top10During['Area']=df['AreaL']+'/'+df['AreaR']+'/'+df['Alarm Msg.']
    df_top10During=df_top10During.groupby(['Area'])['Duration'].agg(['sum'])
    df_top10During=df_top10During.sort_values('sum',ascending=False)
    df_top10During =df_top10During.head(20)



    #############################表连接##############################################
    print(df_PLC)
    print(df_PLCT)
    df_join = pd.concat([df_PLC, df_PLCT], axis=1)
    print(df_join)

    ###############*******     出图，出表     ***###################
    df_PLC2=df_PLC.head(10)
    df_PLCT2=df_PLCT.head(10)
    df_PLC2.to_excel('E:\DATA_ENGIN\\top10'+'\\'+fileName+saveAdress+'T10F.xlsx')
    df_PLCT2.to_excel('E:\DATA_ENGIN\\top10' +'\\'+fileName+saveAdress+'T10T.xlsx')

    ##########################1.柱状图##########################
    def histogram(xvalues,yvalues,a,b,C):
        
        yvalues=np.array(list(yvalues))
        
        plt.bar(xvalues,yvalues)
        title =str(b)+str(a)
        plt.title(title,fontsize =11)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_ticks_position('bottom')
        ax.spines['bottom'].set_position(('data', 0))
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontsize(7)
            label.set_bbox(dict(facecolor = 'lightgreen', edgecolor = 'None', alpha = 0.2))
        plt.xticks(rotation=int(C), fontsize =7)
        plt.subplots_adjust(left=0.18, wspace=0.25, hspace=0.25,bottom=0.36, top=0.91)

        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['savefig.dpi'] = 200  # 图片像素
        plt.rcParams['figure.dpi'] = 200  # 分辨率
        b=b+a
        fig=plt.gcf()
        fig.set_size_inches(9,5) #对图像对象设置大小
        plt.savefig(save+str(b)+'bar.jpg')
        plt.show()
        plt.pause(1) 

        plt.close() 
        print('导出'+str(b)+'bar.jpg')
        
    ###############*******2.. 折线***###################
    import matplotlib.dates as mdate
    
    ####线图#########

    def plot_curve1(data,title):
        fig1 = plt.figure(figsize=(5,5))
        ax1 = fig1.add_subplot(1,1,1)
        ax1.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
        #设置时间标签显示格式
        plt.xticks(list(data.index),rotation=90, fontsize =7) 
        plt.title(title,fontsize =11)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['savefig.dpi'] = 200  # 图片像素
        plt.rcParams['figure.dpi'] = 200  # 分辨率 
        plt.plot(data,'o-')
        plt.subplots_adjust(left=0.18, wspace=0.25, hspace=0.25,bottom=0.26, top=0.91)
        fig=plt.gcf()
        fig.set_size_inches(9,5) #对图像对象设置大小
        
        title=fileName+title
        plt.savefig(save+title+'bar.jpg')   
        plt.show()
        

    def plot_curve(data,title):
    
        # ax1 = fig1.add_subplot(1,1,1)
        # ax1.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
        #设置时间标签显示格式
        plt.xticks(list(data.index),rotation=90, fontsize =7) 
        plt.title(title,fontsize =11)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['savefig.dpi'] = 200  # 图片像素
        plt.rcParams['figure.dpi'] = 200  # 分辨率 
        plt.plot(data,'o-')
        plt.subplots_adjust(left=0.18, wspace=0.25, hspace=0.25,bottom=0.26, top=0.91)
        fig=plt.gcf()
        fig.set_size_inches(9,5) #对图像对象设置大小
        
        title=fileName+title
        plt.savefig(save+title+'bar.jpg')   
        plt.show()
        plt.close(fig)

    def plot_curve2(data1,data2,title):
        # fig1 = plt.figure(figsize=(15,15))
        # ax1 = fig1.add_subplot(1,1,1)
        # ax1.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
        #设置时间标签显示格式
        plt.title(title,fontsize =11)
        plt.xticks(list(data1.index),rotation=90)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['savefig.dpi'] = 200  # 图片像素
        plt.rcParams['figure.dpi'] = 200  # 分辨率 

        plt.plot(data1,'o-',color='Yellow', label='AllFault')
        plt.plot(data2,'o-',color='Red', label='LineStop')
        # plt.plot(data3,'o-')
        # plt.plot(data4,'o-')
        plt.subplots_adjust(left=0.18, wspace=0.25, hspace=0.25,bottom=0.26, top=0.91)
        fig=plt.gcf()
        fig.set_size_inches(9,5) #对图像对象设置大小
        plt.legend(loc="upper right") 
        title=fileName+title
        plt.savefig(save+title+'bar.jpg')

    #######################1. 折线 ################
    """ xvalues =list(df_date_frq.index)
    yvalues = df_date_frq['count']
    histogram(xvalues,yvalues,'每日故障次数',fileName,90) 
    xvalues =list(df_date_during.index)
    yvalues = df_date_during['sum']
    histogram(xvalues,yvalues,'每日故障时间',fileName,90)  """
    plot_curve2(hours,P60,'每小时曲线')


    plot_curve1(df_date_during,'每日故障时间')
    plot_curve1(df_date_frq,'每日故障次数')
    ###############*******2. 柱状***###################

    # xvalues =list(df_message.index)
    # yvalues=df_message['count']
    # plt.show()

    xvalues =list(df_PLCT.index)
    yvalues = df_PLCT['sum']
    histogram(xvalues,yvalues,'故障时长',fileName,90) 
    # ##
    xvalues =list(df_device.index)
    yvalues = df_device['count']
    histogram(xvalues,yvalues,'设备元器件故障次数',fileName,90) 
    # ###
    xvalues =list(df_PLC.index)
    yvalues = df_PLC['count']
    histogram(xvalues,yvalues,'设备条线故障次数',fileName,90) 
    # ##
    xvalues =list(df_message.index)
    yvalues = df_message['count']
    histogram(xvalues,yvalues,'故障报警分型次数',fileName,270) 

    # ####
    xvalues =list(df_type.index)
    yvalues =df_type['count']
    histogram(xvalues,yvalues,'设备种类分型次数',fileName,90)  

    #######
    xvalues= list(df_top10.index)
    yvalues =df_top10['count']
    histogram(xvalues,yvalues,'设备TOP10频次',fileName,270)
    #######
    xvalues= list(df_top10During.index)
    yvalues =df_top10During['sum']
    histogram(xvalues,yvalues,'设备TOP10时长',fileName,270)

    #########################词云
    df_message=df_message.reset_index()

    df_message['Alarm Msg.'].to_csv(fileName+'故障.txt')
    import wordcloud
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    txt = open(fileName+'故障.txt').read()

    wc= wordcloud.WordCloud(width=400,height=400)
    wc =wc.generate(txt)
    plt.imshow(wc, interpolation='bilinear')
    plt.axis('off')

    wc.to_file(save+fileName+'故障.png') 




    ###******************end***####################
    print('完成')