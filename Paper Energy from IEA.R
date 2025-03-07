#Paper Energy from IEA
##运行以下三行代码，保证内存
options(java.parameters='-Xms8g')
gc()###内存不够时跑一下这个
memory.limit(10240000)

install.packages("RJDBC")
memory.limit(102400)
install.packages("openxlsx")
install.packages("xlsx")
install.packages("data.table") 
install.packages("do")
install.packages("installr")
install.packages("stringr")#关于字符的一系列操作



##引用以下包
library(rJava)
library(RJDBC)
library(data.table)
library(xlsxjars)
library(data.table)
library(ggplot2)
library(reshape2)      
library(do)
library(stringr)
library(cowplot)#图表拼接函数
library(RColorBrewer)
library(plyr)
library(basicTrendline)
library(readxl)
library(xlsx)
library(openxlsx)

setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")#设置默认路径


RG=read.xlsx("Input/Energy structure and emissions.xlsx",sheetName = "Nation list")#读入所有国家名称
CI_W=read.xlsx("INPUT/Energy structure and emissions.xlsx",sheetName = "CI1960")#读入能源碳强度数据
CI=melt(CI_W,id=c("Year","Region"))#将原始宽型数据重组为长型数据
names(CI)=c("Year","Region","Type","Value")#更改列名
#折线图,各国造纸行业能源碳强度
p=ggplot(subset(CI,Region!="World"), mapping = aes(x = Year, y = Value, Group=Type,colour=Type)) + geom_line(size=0.8,aes())+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=10,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=8,face="bold"),
        axis.text.y=element_text(size=7,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb", size=0.5, linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green", size=0.5, linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.2, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 8, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 8, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  labs(text=element_text(family='calibri'),y="Carbon intensity: g/MJ",x='Year',title='Carbon intensity of energy consumption',colour="CI")+
  facet_wrap(~Region,ncol=5)
p
ggsave(p, file="OUT/Plot/CI.jpeg", width=16, height=12)

TCV_W=read.xlsx("Input/Energy structure and emissions.xlsx",sheet =  "TCV1960")#输入宽数据，能源标准量
TCV=melt(TCV_W,id=c("Year","Region"))#将原始宽型数据重组为长型数据
names(TCV)=c("Year","Region","Type","Value")#更改列名
row.names(TCV)=1:nrow(TCV)#更改行名

#################################################IEA能源柱状图###################################################################

#筛选数据
#合并类型
TCV_CB_W=read.xlsx("Input/Energy structure and emissions.xlsx",sheet =  "TCV1960_CB")#合并后的数据
TCV_CB_W=TCV_CB_W[,(1:9)]
TCV_CB=melt(TCV_CB_W,id=c("Year","Region"))#将原始宽型数据重组为长型数据
names(TCV_CB)=c("Year","Region","Type","Value")#更改列名

#柱状堆积图，各国能源使用情况
p=ggplot(subset(TCV_CB,Region!="World"),aes(Year,Value/1000,fill=Type),color=enenrgy_plate)+
  geom_col(position="stack",width = 0.6)+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=10,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=12,face="bold"),
        axis.text.y=element_text(size=12,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb", size=0.5, linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green", size=0.5, linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.2, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 8, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 8, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  theme(plot.margin=margin(0.2,0.5,0.2,0.2,"cm"))+####位置顺序为：上、右、下、左
  #theme(legend.position = "none")+
  labs(text=element_text(family='calibri'),y="Energy consumption: PJ",x='Year',title='Energy consumption by PPP Industry')+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_fill_manual(values = c("#07689F","#A6E3E9","#59B791","#FFA8A8","#FD8D14","#7C4789","#A076F9"))
p
ggsave(p, file="OUT/Plot/TCV_CB-231013-图例.jpeg", width=11, height=9)

















#折线图，各国生物质功能
p=ggplot(subset(CI,Region!="World"), mapping = aes(x = Year, y = Value, Group=Type,colour=Type)) + geom_line(size=0.8,aes())+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=10,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=8,face="bold"),
        axis.text.y=element_text(size=7,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb", size=0.5, linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green", size=0.5, linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.2, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 8, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 8, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  labs(text=element_text(family='calibri'),y="Carbon intensity: g/MJ",x='Year',title='Carbon intensity of energy consumption',colour="CI")+
  facet_wrap(~Region,ncol=5)
p
ggsave(p, file="OUT/Plot/CI.jpeg", width=16, height=12)