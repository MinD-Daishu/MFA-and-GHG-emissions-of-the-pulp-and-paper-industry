########GHG Intensity####################
##平均每单位产品的排放强度

##导入包
library(ggplot2)
library(RJDBC)
library(xlsx)
library(reshape2)
library(viridis)

##设置路径
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")


##导入数据
CI=read.xlsx("OUT/Plot/Figures&Tables/Tables.xlsx",sheetName = "GHGIntsty")###打开看一下，去除多余的列
CI=CI[,-(3:4)]
CI[CI<0]=0



# 绘制折线图
p=ggplot(CI, aes(x = Year, y = Intensity)) +
  geom_line(linewidth=0.8) +
  #labs(x = "Year", y = expression(paste("Forest carbon emissions: CO"[2],"-eq"))) +
  theme_bw()+
  theme(plot.margin = margin(0.2, 0.5, 0.2, 0.2, "cm"))+
  theme(legend.position = "none",
    text=element_text(family='calibri'),
        plot.title=element_text(size=18,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=22,face="bold"),
        axis.text.y=element_text(size=22,face="bold"),
        axis.title.x = element_text(size=18,face="bold"),
        axis.title.y = element_text(size=18,face="bold"),
        panel.background = element_blank(),
        #panel.border = element_rect(color = 'black'),
        strip.background.x = element_rect(color="black", fill="#bbbbbb", linetype="solid"),
        panel.spacing.x = unit(0.5, "cm"),
        panel.spacing.y = unit(0.5, "cm"),
        strip.text.x = element_text(size = 18,  face = "bold.italic")
  )+
  theme(panel.grid = element_blank())+
  labs(text=element_text(family='calibri'),y=expression(paste("GHG emissions intensity: kg CO"[2], "-eq per kg product" ),x="Year"))+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_color_manual(values = colors(30))

p

ggsave(p, file="OUT/Plot/GHGemissionsIntensity-free-231016.jpeg", width=20, height=16)


##如果纵坐标各国不固定，那么facet_wrap里面加一个scales="free"