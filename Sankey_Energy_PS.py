# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 17:26:58 2020

@author: zhongguodaishu
"""

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


nodes=[
       {'name':'Energy Input'},
       {'name':'Energy Input (Fossil)'},
       {'name':'Energy Input (Bio)'},
       {'name':'CWP'},
       {'name':'NWP'},
       {'name':'MP'},
       {'name':'RP'},
       {'name':'NP'},
       {'name':'PW'},
       {'name':'HS'},
       {'name':'PP'},
       {'name':'OP'},
       {'name':'PR'},
       ]
data=pd.read_excel( 'Sankey_Energy_PS.xlsx',sheet_name='Energy_PS')
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
         'Energy consumption/EJ',
         nodes,
         links,
         linestyle_opt=opts.LineStyleOpts(opacity=0.35,curve=0.5,color="source"),
         label_opts=opts.LabelOpts(position="right",),
         node_width = 10,
         node_gap = 20,
         layout_iterations= 100
    )
     .set_global_opts(title_opts=opts.TitleOpts(title="Energy_PS"))
     )

pic.render('OUT/Sankey_Energy_PS.html')


