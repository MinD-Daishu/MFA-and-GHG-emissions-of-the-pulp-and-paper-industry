# -*- coding: utf-8 -*-
"""
Created on Sat Dec 26 21:40:32 2020

@author: zhongguodaishu
"""
#导入所需包
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Sankey
#设置工作路径
import os
os.chdir('D:\\坚果云同步文件夹\\PythonBook\\GlobalPaper\\Sankey')
os.getcwd()    #获取当前工作目录

#以下还是尝试，已经成功了，发现问题是本代码无法绘制循环代码
nodes=[
       {'name':'Wood'},
       {'name':'Other fibres'},
       {'name':'Paper for recycling'},
       {'name':'Mechanical pulping'},
       {'name':'Chemical pulping'},
       {'name':'Recycled pulping'},
       {'name':'Mechanical pulp'},
       {'name':'Pulping waste'},
       {'name':'Chemical pulp'},
       {'name':'Recycled pulp'},
       {'name':'Newsprint'},
       {'name':'Printing and writing papers'},
       {'name':'Household and sanitary papers'},
       {'name':'Packaging papers'},
       {'name':'Other papers'},
       {'name':'Use'},
       {'name':'Non-fibrous materials'},
       {'name':'Stocks'},
       {'name':'Recovered paper'},
       {'name':'Incineration'},
       {'name':'Energy recovery'},
       {'name':'Non-energy recovery'},
       {'name':'Landfill'},
       {'name':'Net import'},
       {'name':'Net export'}
       ]
data=pd.read_excel( 'Sankey_China_2019.xlsx')
# 生成links
links = []
for i in data.values:
    dic = {}
    dic['source'] = i[0]
    dic['target'] = i[1]
    dic['value'] = i[2]
    links.append(dic)


pic=(
     Sankey().add(
         'weight/ton',
         nodes,
         links,
         linestyle_opt=opts.LineStyleOpts(opacity=0.35,curve=0.5,color="source"),
         label_opts=opts.LabelOpts(position="right",),
         node_width = 8,
         node_gap = 40,
         layout_iterations= 100
    )
     .set_global_opts(title_opts=opts.TitleOpts(title="China Paper MFA 2019"))
     )

pic.render('OUT/test.html')


#更齐全的节点信息
nodes=[
       {'name':'Wood'},
       {'name':'Other fibres'},
       {'name':'Paper for recycling'},
       {'name':'Mechanical pulping'},
       {'name':'Chemical pulping'},
       {'name':'Recycled pulping'},
       {'name':'Mechanical pulp'},
       {'name':'Pulping waste'},
       {'name':'Chemical pulp'},
       {'name':'Recycled pulp'},
       {'name':'Newsprint'},
       {'name':'Printing and writing papers'},
       {'name':'Household and sanitary papers'},
       {'name':'Packaging papers'},
       {'name':'Other papers'},
       {'name':'Paper for recycling'},
       {'name':'Non-fibrous materials'},
       {'name':'Use'},
       {'name':'Stocks'},
       {'name':'Incineration'},
       {'name':'Energy recovery'},
       {'name':'Non-energy recovery'},
       {'name':'Landfill'},
       {'name':'Net import-MP'},
       {'name':'Net import-CP'},
       {'name':'Net import-RP'},
       {'name':'Net import-RPing'},
       {'name':'Net import-NP'},
       {'name':'Net import-P&W'},
       {'name':'Net import-HS'},
       {'name':'Net import-PP'},
       {'name':'Net import-OP'},
       {'name':'Paper making loss'}       
       ]