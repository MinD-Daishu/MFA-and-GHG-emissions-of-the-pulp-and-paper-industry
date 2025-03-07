# -*- coding: utf-8 -*-
"""
Created on Tue Oct 13 15:21:53 2020

@author: zhongguodaishu
"""
#导入所需包
import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns
#设置工作路径
import os
os.chdir('D:\\坚果云同步文件夹\\PythonBook\\GlobalPaper')
os.getcwd()    #获取当前工作目录

#导入数据
data = pd.read_excel( 'Global production and trade_excludes regions.xlsx' )
df0=pd.DataFrame(data)
df0.fillna(0)
#开始处理数据
dfArea=pd.DataFrame(df0.iloc[:,1])
dfItem=pd.DataFrame(df0.iloc[:,3])
dfElement=pd.DataFrame(df0.iloc[:,5])
dfData=df0.iloc[:,7:66]
df=((dfArea.join(dfItem)).join(dfElement)).join(dfData)
dfT=pd.DataFrame(df.T)
dfTData=dfT.iloc[4:62,:].apply(pd.to_numeric)
dfTString=dfT.iloc[0:3,:]
dfTappend=dfTData.append(dfTString)
dfTappend.columns=df.iloc[:,0]

#报纸生产量#使用原始数据
NpP0=df0.loc[(df0['Item']=='Newsprint') & (df0['Element']=='Production')]#将各国报纸生产量数据筛选出来形成一个数据框
NpPArea=pd.DataFrame(NpP0.iloc[:,1])#生成地区数据框
NpPProduction=NpP0.iloc[:,7:66]#生成生产量数据框
NpP=NpPArea.join(NpPProduction)#将地区和生产量合并
NpPT=NpP.T#为了后面画图，需要转置
NpPT=NpPT.iloc[1:60,:].apply(pd.to_numeric)#转置后发现数据类型都变成了object，所以要转换成数字型
NpPT.columns=NpP.iloc[:,0]#需要把列标签改成地区

 
#绘制折线图，下面的命令要一起运行才能在一页上
plt.plot(NpPTMax20)
plt.title("Newsprint Production")
plt.legend(NpPTMax20.columns,bbox_to_anchor=(0, -1.6), loc=3)
plt.xlabel("Year")
plt.ylabel("Production:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpPTMax20_Lineplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
#绘制堆积图
NpPTMax20.plot.area(stacked=True)
plt.title("Newsprint Production")
plt.legend(NpPTMax20.columns,bbox_to_anchor=(0, -1.6), loc=3)
plt.xlabel("Year")
plt.ylabel("Production:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpPTMax20_Areaplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
#当图例在图外时，用plt.savefig保存的图片图例只保存了一半，
#在plt.savefig中加入“ bbox_inches = 'tight'”参数即可
#plt.savefig('path', bbox_inches = 'tight')
#报纸消费量
Np=df.loc[(df['Item']=='Newsprint') ]#将各国报纸数据筛选出来形成一个数据框
NpP=Np.loc[(Np['Element']=='Production')]#取出生产量数据
NpI=Np.loc[(Np['Element']=='Import Quantity')]#取出进口数据
NpE=Np.loc[(Np['Element']=='Export Quantity')]#取出出口数据
NpP.fillna(0)
NpI.fillna(0)
NpE.fillna(0)
NpC=NpP+NpI-NpE
NpCData=NpP.iloc[:,3:]+NpI.iloc[:,3:]-NpE.iloc[:,3:]#计算表观消费量
DataforPlot=NpC

NpPT=NpP.T#为了后面画图，需要转置
NpPT=NpPT.iloc[1:60,:].apply(pd.to_numeric)#转置后发现数据类型都变成了object，所以要转换成数字型
NpPT.columns=NpP.iloc[:,0]#需要把列标签改成地区
#报纸数据，手工准备好的
NpP = pd.DataFrame(pd.read_excel( 'NewsprintData.xlsx',sheet_name='NpP' ))
NpI = pd.DataFrame(pd.read_excel( 'NewsprintData.xlsx',sheet_name='NpI' ))
NpE = pd.DataFrame(pd.read_excel( 'NewsprintData.xlsx',sheet_name='NpE' ))
NpC = pd.DataFrame(pd.read_excel( 'NewsprintData.xlsx',sheet_name='NpC' ))
###画图专用代码###
#绘制折线图，下面的命令要一起运行才能在一页上
NpPT=NpP.T#为了后面画图，需要转置
NpPT=NpPT.iloc[1:60,:].apply(pd.to_numeric)#转置后发现数据类型都变成了object，所以要转换成数字型
NpPT.columns=NpP.iloc[:,0]#需要把列标签改成地区
DataforPlot=NpPT
plt.plot(DataforPlot)
plt.title("Newsprint Production")
plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -1.6), loc=3)
plt.xlabel("Year")
plt.ylabel("Production:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpC_Lineplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)

#绘制堆积图
DataforPlot.plot.area(stacked=True)
plt.title("Newsprint Production")
plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -1.6), loc=3)
plt.xlabel("Year")
plt.ylabel("Production:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpPTMax20_Areaplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)

