
##腾出内存
options(java.parameters='-Xms8g')
gc()###内存不够时跑一下这个
memory.limit(1024000)

#聚类热力图
#install.packages("pheatmap")


library(pheatmap)
##引用所需包
library(tidyverse)
library(RJDBC)
library(xlsx)
library(data.table)
library(ggplot2)
library(reshape2)
library(do)
library(RColorBrewer)
library(stringr)
library(cowplot)#图表拼接函数
library(RColorBrewer)

##设置路径
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper/OUT")
##导入数据
SC_SA=read.xlsx("Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "HeatMap")
SC_SA=SC_SA[1:30,1:17]
SC_SA[,2:17]=round(SC_SA[,2:17],2)
SC_SA_m=as.matrix(SC_SA)
SC_SA_m=SC_SA_m[,-1]
SC_SA_m=apply(SC_SA_m,2,as.numeric)
rownames(SC_SA_m)=SC_SA$Region
SC_SA_m_t=t(SC_SA_m)##转置
data_pct=SC_SA_m_t




##色条设置
#breaks
bk1 <- c(seq(-2,-0.01,by=0.01),seq(0,2,by=0.01))
#breaks
bk2 <- c(seq(-200,-10,by=1),seq(0,200,by=1))

##配色1（暗红色+暗蓝色）
#colorRampPalette(colors = c("#003399","#d8e9ef"))(length(bk1)/2),
#colorRampPalette(colors = c("white","#b41b09"))(length(bk1)/2))
##配色2（大红+大蓝）
#colorRampPalette(colors = c("#02359A","white"))(length(bk1)/2),
#colorRampPalette(colors = c("white","red"))(length(bk1)/2))
##配色3（红白蓝绿）
#colorRampPalette(colors = c ("#83BD75","#E9EFC0"))(length(bk1)/4),
#colorRampPalette(colors = c("#E9EFC0","#A2DCF7"))(length(bk1)/4),
#colorRampPalette(colors = c("white","red"))(length(bk1)/2))
##配色4（红白蓝黄绿）
#colorRampPalette(colors = c ("#83BD75","#E9EFC0"))(length(bk1)*101/401),
#colorRampPalette(colors = c("#4DC1EA","#EFFFFD"))(length(bk1)*99/401),
#colorRampPalette(colors = c("white"))(length(bk1)1/401),
#colorRampPalette(colors = c("#FFEDED","red"))(length(bk1)*200/401))


p1=pheatmap(data_pct, cluster_row = F,cluster_col = TRUE,   border_color = "black",
            display_numbers=TRUE, number_format = "%.1f",number_color = "black",
            cellwidth = 16, cellheight = 16, fontsize_number = 8,fontsize_row=10,fontsize_col=10, breaks = bk1,
            annotation_row_names_position="right",
            color=c(colorRampPalette(colors = c ("#83BD75","#E9EFC0"))(length(bk1)*100/401),
              colorRampPalette(colors = c("#4DC1EA","#EFFFFD"))(length(bk1)*100/401),
              colorRampPalette(colors = c("white"))(length(bk1)/401),
              colorRampPalette(colors = c("#FFEDED","red"))(length(bk1)*200/401)),
            legend_breaks=seq(-1,1,1),
            angle_col=90,filename = "Plot/HeatMap/SC_SM_heatmap_color4-230922-显示聚类图.png")

##以下一步操作，做聚类的同时，不显示聚类图
y=rownames(data_pct)
x=colnames(data_pct)[p1$tree_col[["order"]]]
new_data_pct=data_pct[y,x]
p2=pheatmap(new_data_pct, cluster_row = FALSE,cluster_col = FALSE, border_color = "black",
            display_numbers=TRUE, number_format = "%.1f",number_color = "black",
            cellwidth = 17, cellheight = 17, fontsize_number = 7.5,fontsize_row=10,fontsize_col=10, breaks = bk1,
            annotation_row_names_position="right",
            color=c(colorRampPalette(colors = c ("#83BD75","#E9EFC0"))(length(bk1)*100/401),
                    colorRampPalette(colors = c("#4DC1EA","#EFFFFD"))(length(bk1)*100/401),
                    colorRampPalette(colors = c("white"))(length(bk1)/401),
                    colorRampPalette(colors = c("#FFEDED","red"))(length(bk1)*200/401)),
            legend_breaks=seq(-1,1,1),
            angle_col=90,filename = "Plot/HeatMap/SC_SM_heatmap_color4-231106-去掉聚类图.pdf")




###########绘制单措施情景的绝对值热力图###############



##导入数据
SC_SA=read.xlsx("Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "HeatMap_Original")
SC_SA=SC_SA[1:30,1:19]
SC_SA[,2:19]=round(SC_SA[,2:19],2)
SC_SA_m=as.matrix(SC_SA)
SC_SA_m=SC_SA_m[,-1]
SC_SA_m=apply(SC_SA_m,2,as.numeric)
rownames(SC_SA_m)=SC_SA$Region
SC_SA_m_t=t(SC_SA_m)##转置
data_org=SC_SA_m_t
##色条设置
#breaks
bk3 <- c(seq(-100,-0.02,by=0.01),seq(0,300,by=0.01))



p3=pheatmap(data_org, display_numbers=TRUE, number_format = "%.0f",number_color = "blue",
            cluster_row = F,cluster_col = T,border_color = "black",
            cellwidth = 15, cellheight = 15, fontsize_number = 8,fontsize_row=10,fontsize_col=10,breaks = bk3,
            color=c(colorRampPalette(colors = c("#79BEDB"))(length(bk3)*1/4),
                    colorRampPalette(colors = c("white","red"))(length(bk3)*3/4)),
            legend_breaks=seq(-100,300,100),
            angle_col=90,filename = "Plot/HeatMap/SC_SA_heatmap_original-0922-显示聚类图.png")
##以下一步操作，做聚类的同时，不显示聚类图
y=rownames(data_org)
x1=colnames(data_pct)[p1$tree_col[["order"]]]
x3=colnames(data_org)[p3$tree_col[["order"]]]
p1_data_org=data_org[y,x1]
p3_data_org=data_org[y,x3]

p4=pheatmap(p1_data_org, display_numbers=TRUE, number_format = "%.0f",number_color = "blue",
            cluster_row = F,cluster_col = F,border_color = "black",
            cellwidth = 15, cellheight = 15, fontsize_number = 8,fontsize_row=9,fontsize_col=10,breaks = bk3,
            color=c(colorRampPalette(colors = c("#79BEDB"))(length(bk3)*9999/40001),
                    colorRampPalette(colors = c("white","red"))(length(bk3)*30002/40001)),
            legend_breaks=seq(-100,300,100),
            angle_col=90,filename = "Plot/HeatMap/SC_SA_heatmap_original-230922-去掉聚类图-正文顺序.png")
p5=pheatmap(p3_data_org, display_numbers=TRUE, number_format = "%.0f",number_color = "blue",
            cluster_row = F,cluster_col = F,border_color = "black",
            cellwidth = 15, cellheight = 15, fontsize_number = 8,fontsize_row=9,fontsize_col=10,breaks = bk3,
            color=c(colorRampPalette(colors = c("#79BEDB"))(length(bk3)*9999/40001),
                    colorRampPalette(colors = c("white","red"))(length(bk3)*30002/40001)),
            legend_breaks=seq(-100,300,100),
            angle_col=90,filename = "Plot/HeatMap/SC_SA_heatmap_original-230922-去掉聚类图.png")

###这里要输出两张图，分别是按照正文里热力图国家顺序的图p4，和按照绝对值聚类后顺序的图p5。
###第一个图，要输入p1的国家顺序
###第二个图，直接聚类输出即可
###p3只用来输出聚类后且包含聚类图的图，暂时用不到，只是放这里



