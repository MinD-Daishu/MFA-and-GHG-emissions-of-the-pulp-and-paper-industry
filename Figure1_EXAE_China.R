############## 图1 part a 全球分阶段的累积图 #######################
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

##########################增加一个调色板#############################
library(RColorBrewer)
colourCount = length(unique(mtcars$hp))
getPalette1 = colorRampPalette(brewer.pal(9, "Paired"))

ggplot(mtcars) + 
  geom_bar(aes(factor(hp)), fill=getPalette1(colourCount)) + 
  theme(legend.position="right")
######################## 调色板2 #############################
colName = c( "#FF6600","#f9b271","#F78500",#碳存储量
            "#adbd37","#588133","#00a03e",#采集阶段
            "#fc9d9a",#化学木浆，红1
            "#FF717E",#机械木浆MP，红2
            "#FF3B1D",#RP，红3
            "#82052f",#非木浆，红4
            "#d8e9ef",#包装纸PP，蓝1
            "#66c2cb",#PW,蓝2
            "#42B4E7" ,#报纸，蓝3
            "#6195C5",#卫生纸，蓝4
            "#005B9A",#其他纸张，蓝5
            "#8C489F",#PR，紫色
            "#663300",#废水
            "#C04D00",#填埋
            "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下

######################## 调色板3 #############################
##满足Nature不允许使用彩虹色的要求，把颜色岔开
colName = c( "#BEADFA","#A2C579","#D8B384",
             "#F3C1C6","#C1D0B5","#FF9292" ,
             "#FBF0B2","#8CC0DE","#FFD9C0",
             "#ACBCFF","#B2B8A3","#C67ACE",
             "#73A9AD",
             "#FAAB78",
             "#A3A847",
             "#967E76",
             "#FF6969",
             "#1C6DD0",
             "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下


######################## 调色板4 #############################
##采用colorbrewer2，色盲友好、印刷友好色彩
colName = c( "#756bb1","#bcbddc","#efedf5",#碳存储量
             "#adbd37","#addd8e","#31a354",#采集阶段
             "#feebe2",#非木浆，红4
             "#fbb4b9",#RP，红3
             "#f768a1",#机械木浆MP，红2
             "#ae017e",#化学木浆，红1
             "#ffffcc",#包装纸PP，蓝1
             "#a1dab4",#PW,蓝2
             "#41b6c4" ,#报纸，蓝3
             "#2c7fb8",#卫生纸，蓝4
             "#253494",#其他纸张，蓝5
             "#fc9272",#PR，紫色
             "#bdbdbd", #废水
             "#636363",#填埋
             "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下

######################## 调色板5 【最终选用】#############################
colName = c( "#54278f","#756bb1","#bcbddc",#碳存储量
             "#adbd37","#588133","#004225",#采集阶段
             "#fee6ce",#化学木浆，红1
             "#fdae6b",#机械木浆MP，红2
             "#f16913",#RP，红3
             "#a63603",#非木浆，红4
             "#d8e9ef",#包装纸PP，蓝1
             "#9ecae1",#PW,蓝2
             "#6baed6" ,#报纸，蓝3
             "#3182bd",#卫生纸，蓝4
             "#08519c",#其他纸张，蓝5
             "#08306b",#PR，紫色
             "#850000",#废水
             "#E26868",#填埋
             "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下

######################## 调色板6 #############################
colName = c( "#bcbddc","#756bb1","#54278f",#碳存储量
             "#adbd37","#588133","#004225",#采集阶段
             "#fee6ce",#化学木浆，红1
             "#fdae6b",#机械木浆MP，红2
             "#f16913",#RP，红3
             "#a63603",#非木浆，红4
             "#d8e9ef",#包装纸PP，蓝1
             "#9ecae1",#PW,蓝2
             "#6baed6" ,#报纸，蓝3
             "#3182bd",#卫生纸，蓝4
             "#08519c",#其他纸张，蓝5
             "#08306b",#PR，紫色
             "#E26868",#废水
             "#850000",#填埋
             "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下

######################## 调色板7 #############################
colName = c( "#bcbddc","#756bb1","#54278f",#碳存储量
             "#adbd37","#588133","#004225",#采集阶段
             "#fee6ce",#化学木浆，红1
             "#fdae6b",#机械木浆MP，红2
             "#f16913",#RP，红3
             "#a63603",#非木浆，红4
             "#d8e9ef",#包装纸PP，蓝1
             "#9ecae1",#PW,蓝2
             "#6baed6" ,#报纸，蓝3
             "#3182bd",#卫生纸，蓝4
             "#08519c",#其他纸张，蓝5
             "#08306b",#PR，紫色
             "#850000",#废水
             "#E26868",#填埋
             "white"#NCE
)#无限制，随便添
colors = colorRampPalette(colName)#N为需要的颜色数量
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#展示一下


##########################碳排+碳汇（不包含生物质）############################
##准备数据
Global_Carbon=read.xlsx("OUT/China_PPI_GHG.xlsx",sheetName = "CE")
Global_Carbon_L=melt(Global_Carbon,id=c("Year"),variable.name = "Process",value.name = "Value")#将原始宽型数据重组为长型数据
Global_Carbon_stage=subset(Global_Carbon_L, Process!="uNCE"&Process!="vNCE_EXAE"&Process!="aAE_lf"&Process!="bAE_er"
                          )
Global_Carbon_total=subset(Global_Carbon_L, Process=="vNCE_EXAE")

p=ggplot(Global_Carbon_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(Global_Carbon_total,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8)+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=14,face="bold"),
        axis.text.y=element_text(size=14,face="bold"),
        axis.title.x = element_text(size=12,face="bold"),
        axis.title.y = element_text(size=12,face="bold"))+
  theme( strip.background.x = element_rect(
    color="black", fill="#bbbbbb", size=0.5, linetype="solid"
  ),
  strip.background.y = element_rect(
    color="black", fill="green", size=0.5, linetype="solid"
  ),
  panel.border = element_rect(color = 'black'))+
  theme(panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"))+
  theme(
    strip.text.x = element_text(
      size = 14, color = "black", face = "bold.italic"
    ), # 这里设置x轴方向的字体类型，
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # 这里设置y轴方向的字体类型，
  )+
  labs(text=element_text(family='calibri'),y="CO2 emission: Mt CO2-eq",x="Year")+
  scale_fill_manual(values = colors(19))+
  guides(fill=guide_legend(ncol=2)) 
p
ggsave(p, file="OUT/Plot/Fig1_ChinaCarbon_BCN_EXAE250217.svg", width=8, height=5)

