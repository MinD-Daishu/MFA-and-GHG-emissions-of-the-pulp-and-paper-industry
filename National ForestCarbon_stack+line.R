#Paper Plot
options(java.parameters='-Xms6144m')
memory.limit(102400)
install.packages("RJDBC")
install.packages("openxlsx")
install.packages("xlsx")
install.packages("data.table") 
install.packages("do")
install.packages("installr")
install.packages("stringr")#关于字符的一系列操作
install.packages("cowplot")
install.packages("RColorBrewer")
install.packages("maps")
install.packages("rworldmap")
install.packages("ggplot2")
install.packages("MAP")
install.packages('ggsci')

library(ggsci)
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
library(maps)
library(rworldmap)
library(MAP)
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")


##########################【英文版-百万吨】森林碳排放############################
##准备数据
FC=read.xlsx("Input/ForestCarbon/ForestCarbon_Info.xlsx",sheetName = "CE_df_Brkdn_L")
FC=FC[,1:6]
data_m=melt(FC,id=c("NewID","Region","Year"),variable.name = "Process",value.name = "Value")#将原始宽型数据重组为长型数据
data_m$Year=as.numeric(data_m$Year)

colName = c( "#6BCB77","#FEEB97","#79B8D1"
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:3,y=1,z=as.matrix(1:3),col=colors(3))#展示一下

#####再画总体的30国,按照重新排列的顺序#################
data_stage=subset(data_m, Process!="f_Net.emissions")
data_total=subset(data_m, Process=="f_Net.emissions")
data_stage_neworder=data_stage
data_stage_neworder=data_stage_neworder[order(data_stage_neworder$NewID), ]

data_total_neworder=data_total
data_total_neworder=data_total_neworder[order(data_total_neworder$NewID), ]

p=ggplot(data_stage_neworder,aes(Year,as.numeric(Value)/10^6,fill=Process))+
  geom_area(position="stack")+
  #geom_line(data_total_neworder,mapping=aes(x=Year,y=as.numeric(Value)/10^6),size=0.8,linetype="dashed")+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=18,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=18,face="bold"),
        axis.text.y=element_text(size=18,face="bold"),
        axis.title.x = element_text(size=18,face="bold"),
        axis.title.y = element_text(size=18,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb",  linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green",linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 18, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 18, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  theme(legend.position = "none")+
  labs(text=element_text(family='calibri'),
       y=expression(paste("Forest carbon emissions: million tons CO"[2], "-eq")),
       x="Year")+
  theme(plot.margin = margin(0.2, 1, 0.2, 0.2, "cm"))+
  facet_wrap(~NewID,ncol=5,scales="free")+
  scale_fill_manual(values = colors(3))
p
ggsave(p, file="OUT/Plot/ForestCarbon/ForestCarbon_stack_line_NewOrder.jpeg", width=18, height=6)

##########################【中文版-亿吨】森林碳排放############################
##准备数据
FC=read.xlsx("Input/ForestCarbon/ForestCarbon_Info.xlsx",sheetName = "CE_df_Brkdn_L")
FC=FC[,1:6]
data_m=melt(FC,id=c("NewID","Region","Year"),variable.name = "Process",value.name = "Value")#将原始宽型数据重组为长型数据
data_m$Year=as.numeric(data_m$Year)

colName = c( "#6BCB77","#FEEB97","#79B8D1"
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:3,y=1,z=as.matrix(1:3),col=colors(3))#展示一下

#####再画总体的30国,按照重新排列的顺序#################
data_stage=subset(data_m, Process!="f_Net.emissions")
data_total=subset(data_m, Process=="f_Net.emissions")
data_stage_neworder=data_stage
data_stage_neworder=data_stage_neworder[order(data_stage_neworder$NewID), ]

data_total_neworder=data_total
data_total_neworder=data_total_neworder[order(data_total_neworder$NewID), ]

p=ggplot(data_stage_neworder,aes(Year,as.numeric(Value)/10^8,fill=Process))+
  geom_area(position="stack")+
  #geom_line(data_total_neworder,mapping=aes(x=Year,y=as.numeric(Value)/10^6),size=0.8,linetype="dashed")+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=18,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=18,face="bold"),
        axis.text.y=element_text(size=18,face="bold"),
        axis.title.x = element_text(size=18,face="bold"),
        axis.title.y = element_text(size=18,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb",  linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green",linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 18, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 18, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  theme(legend.position = "none")+
  labs(text=element_text(family='calibri'),
       y=expression(paste("Forest carbon emissions: million tons CO"[2], "-eq")),
       x="Year")+
  theme(plot.margin = margin(0.2, 1, 0.2, 0.2, "cm"))+
  facet_wrap(~NewID,ncol=5,scales="free")+
  scale_fill_manual(values = colors(3))+
  scale_y_continuous(labels = scales::number_format(accuracy = 0.01))
p
ggsave(p, file="OUT/Plot/ForestCarbon/ForestCarbon_stack_line_NewOrder_亿吨.jpeg", width=18, height=6)

