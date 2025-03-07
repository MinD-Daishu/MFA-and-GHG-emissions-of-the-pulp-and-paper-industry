########ForestCarbon_Sen####################
##森林碳敏感性分析（针对认证比例）

##导入包
library(ggplot2)
library(RJDBC)
library(xlsx)
library(reshape2)

##设置路径
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")

##导入数据
FC=read.xlsx("Input/ForestCarbon/ForestCarbon_Info.xlsx",sheetName = "CE_df_Brkdn_L_Sstvty")###打开看一下，去除多余的列
FC=FC[,1:6]

# 绘制折线图
p=ggplot(FC, aes(x = Year, y = FCE_SFC/10^6)) +
  geom_line() +
  #labs(x = "Year", y = expression(paste("Forest carbon emissions: CO"[2],"-eq"))) +
  geom_ribbon(aes(ymin = FCE_SFC/10^6, ymax = FCE_Low1/10^6), fill = "pink", alpha = 0.5)+
  geom_ribbon(aes(ymin = FCE_Low1/10^6, ymax = FCE_Low2/10^6), fill = "lightblue", alpha = 0.5) +
  theme_bw()+
  theme(plot.margin = margin(0.2, 0.5, 0.2, 0.2, "cm"))+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=18,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=18,face="bold"),
        axis.text.y=element_text(size=18,face="bold"),
        axis.title.x = element_text(size=18,face="bold"),
        axis.title.y = element_text(size=18,face="bold"),
        panel.background = element_blank(),
        #panel.border = element_rect(color = 'black'),
        strip.background.x = element_rect(color="black", fill="#bbbbbb", linetype="solid"),
        panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"),
        strip.text.x = element_text(size = 18,  face = "bold.italic")
  )+
  #theme(panel.grid = element_blank())+
  labs(text=element_text(family='calibri'),y=expression(paste("Forest carbon emissions: million tons CO"[2], "-eq"),x="Year"))+
  facet_wrap(~NewID,ncol=5,scales="free")

p

ggsave(p, file="OUT/Plot/ForestCarbon/ForestCarbon_Sens.jpeg", width=18, height=6)




###########################中文版本-亿吨###########################
##导入数据
FC=read.xlsx("Input/ForestCarbon/ForestCarbon_Info.xlsx",sheetName = "CE_df_Brkdn_L_Sstvty")###打开看一下，去除多余的列
FC=FC[,1:6]

# 绘制折线图
p=ggplot(FC, aes(x = Year, y = FCE_SFC/10^8)) +
  geom_line() +
  #labs(x = "Year", y = expression(paste("Forest carbon emissions: CO"[2],"-eq"))) +
  geom_ribbon(aes(ymin = FCE_SFC/10^8, ymax = FCE_Low1/10^8), fill = "pink", alpha = 0.5)+
  geom_ribbon(aes(ymin = FCE_Low1/10^8, ymax = FCE_Low2/10^8), fill = "lightblue", alpha = 0.5) +
  theme_bw()+
  theme(plot.margin = margin(0.2, 0.5, 0.2, 0.2, "cm"))+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=18,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=18,face="bold"),
        axis.text.y=element_text(size=18,face="bold"),
        axis.title.x = element_text(size=18,face="bold"),
        axis.title.y = element_text(size=18,face="bold"),
        panel.background = element_blank(),
        #panel.border = element_rect(color = 'black'),
        strip.background.x = element_rect(color="black", fill="#bbbbbb", linetype="solid"),
        panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"),
        strip.text.x = element_text(size = 18,  face = "bold.italic")
        )+
  #theme(panel.grid = element_blank())+
  labs(text=element_text(family='calibri'),y=expression(paste("Forest carbon emissions: million tons CO"[2], "-eq"),x="Year"))+
  facet_wrap(~NewID,ncol=5,scales="free")+
  scale_y_continuous(labels = scales::number_format(accuracy = 0.01))

p

ggsave(p, file="OUT/Plot/ForestCarbon/ForestCarbon_Sens_亿吨.jpeg", width=18, height=6)
