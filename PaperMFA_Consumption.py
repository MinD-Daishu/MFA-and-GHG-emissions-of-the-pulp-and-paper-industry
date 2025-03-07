# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 14:47:56 2020

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

#用循环解决重复问题
#构建遍历产品名单
Product=['Pulp for paper','Pulp from fibres other than wood','Recovered fibre pulp','Wood pulp',
         'Chemical wood pulp','Mechanical and semi-chemical wood pulp','Mechanical wood pulp',
         'Newsprint','Household and sanitary papers','Printing and writing papers',
         'Graphic papers','Recovered paper','Wrapping papers','Other paper and paperboard',
         'Packaging paper and paper board','Paper and paperboard']
for i in Product:
    Item=df0.loc[(df0['Item']==i)]
    ItemP=Item.loc[(Item['Element']=='Production')]
    ItemI=Item.loc[(Item['Element']=='Import Quantity')]
    ItemE=Item.loc[(Item['Element']=='Export Quantity')]
    ItemP=pd.merge(AreaName,ItemP,on=['Area'],how='left')#合并地区名称和数据
    ItemI=pd.merge(AreaName,ItemI,on=['Area'],how='left')#合并地区名称和数据
    ItemE=pd.merge(AreaName,ItemE,on=['Area'],how='left')#合并地区名称和数据
    del ItemP['Item']
    del ItemP['Element']
    del ItemI['Item']
    del ItemI['Element']
    del ItemE['Item']
    del ItemE['Element']
    
    ItemC=ItemP+ItemI-ItemE
    ItemC=ItemC.fillna(0)
    ItemCTobject=ItemC.T
    ItemCT=ItemCTobject.iloc[1:60,:].apply(pd.to_numeric)
    ItemCT.columns=ItemC.iloc[:,0]#需要把列标签改成地区
    writer=pd.ExcelWriter('OUT20201026/'+i+'_P'+'.xlsx')
    ItemCT.to_excel(writer,sheet_name=i[0:29]+'_P')
    writer.save()
    #按照2019年数据排序
    ItemCsort=ItemCT.sort_values(by=2019,axis=1,ascending=False, inplace=False, na_position='last')
    #选前50名
    ItemCMax50=ItemCsort.iloc[:,0:50]
    ###画图专用代码###
    #绘制折线图，下面的命令要一起运行才能在一页上
    DataforPlot=ItemCMax50
    plt.plot(DataforPlot)
    plt.title(i+" Consumption")
    plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
    plt.xlabel("Year")
    plt.ylabel("Consumption"+":tons")
    plt.margins(0.03,0.03,tight=True)
    plt.savefig("OUT20201026/Figures20201026/"+i+"Consumption"+"Max50_Lineplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
    plt.clf()#添加上这一行，画完第一个图后，重置一下
    #绘制堆积图
    DataforPlot.plot.area(stacked=True)
    plt.title(i+" Consumption")
    plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
    plt.xlabel("Year")
    plt.ylabel("Consumption"+":tons")
    plt.margins(0.03,0.03,tight=True)
    plt.savefig("OUT20201026/Figures20201026/"+i+"Consumption"+"Max50_Areaplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
    plt.clf()#添加上这一行，画完第一个图后，重置一下
    
#测试
Product=['Pulp for paper','Pulp from fibres other than wood','Recovered fibre pulp','Wood pulp',
         'Chemical wood pulp','Mechanical and semi-chemical wood pulp','Mechanical wood pulp',
         'Newsprint','Household and sanitary papers','Printing and writing papers',
         'Graphic papers','Recovered paper','Wrapping papers','Other paper and paperboard',
         'Packaging paper and paper board','Paper and paperboard']
for i in Product:
    Item=df0.loc[(df0['Item']==i)]
    ItemP=Item.loc[(Item['Element']=='Production')]
    ItemI=Item.loc[(Item['Element']=='Import Quantity')]
    ItemE=Item.loc[(Item['Element']=='Export Quantity')]
    ItemP=pd.merge(AreaName,ItemP,on=['Area'],how='left')#合并地区名称和数据
    ItemI=pd.merge(AreaName,ItemI,on=['Area'],how='left')#合并地区名称和数据
    ItemE=pd.merge(AreaName,ItemE,on=['Area'],how='left')#合并地区名称和数据
    del ItemP['Item']
    del ItemP['Element']
    del ItemI['Item']
    del ItemI['Element']
    del ItemE['Item']
    del ItemE['Element']
    ItemPN=ItemP.iloc[:,1:60].apply(pd.to_numeric)
    ItemIN=ItemI.iloc[:,1:60].apply(pd.to_numeric)
    ItemEN=ItemE.iloc[:,1:60].apply(pd.to_numeric)
    ItemC=ItemPN+ItemIN-ItemEN
    ItemC=ItemC.fillna(0)
    ItemCTobject=ItemC.T
    ItemCT=ItemCTobject.iloc[1:60,:].apply(pd.to_numeric)
    ItemCT.columns=ItemP.iloc[:,0]#需要把列标签改成地区
    writer=pd.ExcelWriter('OUT_Consumption/'+i+'_C'+'.xlsx')
    ItemCT.to_excel(writer,sheet_name=i[0:29]+'_C')
    writer.save()
    #按照2019年数据排序
    ItemCsort=ItemCT.sort_values(by=2019,axis=1,ascending=False, inplace=False, na_position='last')
    #选前50名
    ItemCMax50=ItemCsort.iloc[:,0:50]
    ###画图专用代码###
    #绘制折线图，下面的命令要一起运行才能在一页上
    ItemCMax50[ItemCMax50 <0] = 0
    DataforPlot=ItemCMax50
    plt.plot(DataforPlot)
    plt.title(i+" Consumption")
    plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
    plt.xlabel("Year")
    plt.ylabel("Consumption"+":tons")
    plt.margins(0.03,0.03,tight=True)
    plt.savefig("OUT_Consumption/Figures_Consumption/"+i+"Consumption"+"Max50_Lineplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
    plt.clf()#添加上这一行，画完第一个图后，重置一下
    #绘制堆积图
    DataforPlot.plot.area(stacked=True)
    plt.title(i+" Consumption")
    plt.legend(DataforPlot.columns,bbox_to_anchor=(0, -0.2), loc=2)
    plt.xlabel("Year")
    plt.ylabel("Consumption"+":tons")
    plt.margins(0.03,0.03,tight=True)
    plt.savefig("OUT_Consumption/Figures_Consumption/"+i+"Consumption"+"Max50_Areaplot.jpg",dpi=600,bbox_inches = 'tight', pad_inches = 0.5)
    plt.clf()#添加上这一行，画完第一个图后，重置一下    
        
