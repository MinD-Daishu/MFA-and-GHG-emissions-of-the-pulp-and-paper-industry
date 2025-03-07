# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 17:25:31 2020

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
data = pd.read_excel( 'Global production and trade_excludes regions.xlsx',sheet_name='Forestry_All Data' )
df0=pd.DataFrame(data)
df0.fillna(0)
#导入地区
AreaName=pd.read_excel('Global production and trade_excludes regions.xlsx',sheet_name='Nations')

#以下作为一个计算和画图范例，为后面的循环探索做准备
#报纸数据
Np=df0.loc[(df0['Item']=='Newsprint')]#将各国报纸数据筛选出来形成一个数据框
NpP=Np.loc[(Np['Element']=='Production')]#取出生产量数据
NpP=pd.merge(AreaName,NpP,on=['Area'],how='left')#合并地区名称和数据
del NpP['Item']
del NpP['Element']
NpI=Np.loc[(Np['Element']=='Import Quantity')]#取出进口数据
NpI=pd.merge(AreaName,NpI,on=['Area'],how='left')#合并地区名称和数据
del NpI['Item']
del NpI['Element']
NpE=Np.loc[(Np['Element']=='Export Quantity')]#取出出口数据
NpE=pd.merge(AreaName,NpE,on=['Area'],how='left')#合并地区名称和数据
del NpE['Item']
del NpE['Element']
NpP=NpP.fillna(0)
NpI=NpI.fillna(0)
NpE=NpE.fillna(0)

#计算消费量
NpEminus=NpE*(-1)
NpEminus.fillna(0)
NpC=(NpP.add(NpI)).add(NpEminus)
NpC['Area']=NpP['Area']
NpCTobject=NpC.T
NpCT=NpCTobject.iloc[1:60,:].apply(pd.to_numeric)
NpCT.columns=NpC.iloc[:,0]#需要把列标签改成地区

#新建一个表格工作簿
writer=pd.ExcelWriter('OUT/Newsprint2.xlsx')
#把上述数据写入
NpP.to_excel(writer,sheet_name='Newsprint Production')
NpI.to_excel(writer,sheet_name='Newsprint Import')
NpE.to_excel(writer,sheet_name='Newsprint Export')
NpCT.to_excel(writer,sheet_name='Newsprint Consumption')
NpCT[NpCT<0]=0#去除负值
NpCT.to_excel(writer,sheet_name='Newsprint Consumption_rmv')#保存一个去除负值的文件
#保存到本地
writer.save()
#按照2019年数据排序
NpCTsort=NpCT.sort_values(by=2019,axis=1,ascending=False, inplace=False, na_position='last')
#选前20名
NpCTMax50=NpCTsort.iloc[:,0:50]

#为批量出图做准备

###画图专用代码###
#绘制折线图，下面的命令要一起运行才能在一页上
DataforPlot=NpCTMax50
plt.plot(DataforPlot)
plt.title("Newsprint Consumption")
plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
plt.xlabel("Year")
plt.ylabel("Consumption:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpCTMax50_Lineplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
#绘制堆积图
DataforPlot.plot.area(stacked=True)
plt.title("Newsprint Consumption")
plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
plt.xlabel("Year")
plt.ylabel("Consumption:tons")
plt.margins(0.03,0.03,tight=True)
plt.savefig("OUT/Figures/NpCTMax50_Areaplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)





