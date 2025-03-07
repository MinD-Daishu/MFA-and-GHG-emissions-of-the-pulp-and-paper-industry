#Paper Plot
options(java.parameters='-Xms6144m')
memory.limit(102400)
install.packages("RJDBC")
install.packages("openxlsx")
install.packages("xlsx")
install.packages("data.table") 
install.packages("do")
install.packages("installr")
install.packages("stringr")#?????Ö·???Ò»Ïµ?Ð²???
install.packages("cowplot")
install.packages("RColorBrewer")
install.packages("maps")
install.packages("rworldmap")
install.packages("ggplot2")
install.packages("MAP")
install.packages("svglite")

library(RJDBC)
library(xlsx)
library(data.table)
library(ggplot2)
library(reshape2)
library(do)
library(RColorBrewer)
library(stringr)
library(cowplot)#Í¼??Æ´?Óº???
library(RColorBrewer)
library(maps)
library(rworldmap)
library(MAP)
library(svglite)
setwd("D:/??????Í¬???Ä¼???/RBOOK/Global Paper")


##########################???Åµ??????ï¡¿????+Ì¼???????????????Ê£?-???????????Å·?############################
##×¼??????
CarbonAll=read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "Carbon_BCN_PS")
CarbonAll=CarbonAll[1:1770,1:25]
data_m=melt(CarbonAll,id=c("NewID","Region","Year"),variable.name = "Process",value.name = "Value")#??Ô­Ê¼????????????Îª????????
data_m$Year=as.numeric(data_m$Year)
data_stage=subset(data_m, Process!="uNCE"&Process!="vNCE_EXAE"&Process!="aAE_lf"&Process!="bAE_er")
data_total=subset(data_m, Process=="vNCE_EXAE")
colName = c( "#f67504","#fecbaa","#feb64a",#Ì¼?æ´¢??,?????Å·?
             "#adbd37","#588133","#00a03e",#?É¼??×¶?
             "#fc9d9a",#??Ñ§Ä¾??????1
             "#FF717E",#??ÐµÄ¾??MP????2
             "#FF3B1D",#RP????3
             "#82052f",#??Ä¾??????4
             "#d8e9ef",#??×°Ö½PP??À¶1
             "#66c2cb",#PW,À¶2
             "#42B4E7" ,#??Ö½??À¶3
             "#6195C5",#????Ö½??À¶4
             "#005B9A",#????Ö½?Å£?À¶5
             "#8C489F",#PR????É«
             "#663300",#??Ë®
             "#C04D00",#????
             "white"#NCE
)#?????Æ£???????
colors = colorRampPalette(colName)#NÎª??Òª????É«??Á¿
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#Õ¹Ê¾Ò»??

######################## ??É«??5 ??????Ñ¡?Ã¡?#############################
colName = c( "#54278f","#756bb1","#bcbddc",#Ì¼?æ´¢Á¿
             "#adbd37","#588133","#004225",#?É¼??×¶?
             "#fee6ce",#??Ñ§Ä¾??????1
             "#fdae6b",#??ÐµÄ¾??MP????2
             "#f16913",#RP????3
             "#a63603",#??Ä¾??????4
             "#d8e9ef",#??×°Ö½PP??À¶1
             "#9ecae1",#PW,À¶2
             "#6baed6" ,#??Ö½??À¶3
             "#3182bd",#????Ö½??À¶4
             "#08519c",#????Ö½?Å£?À¶5
             "#08306b",#PR????É«
             "#850000",#??Ë®
             "#E26868",#????
             "white"#NCE
)#?????Æ£???????
colors = colorRampPalette(colName)#NÎª??Òª????É«??Á¿
image(x=1:19,y=1,z=as.matrix(1:19),col=colors(19))#Õ¹Ê¾Ò»??



##?È»???????30??,??????Ä¸Ë³??
p=ggplot(data_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(data_total,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=12,face="bold"),
        axis.text.y=element_text(size=12,face="bold"),
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  labs(text=element_text(family='calibri'),y="CO2 emission: Mt CO2-eq",x="Year")+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_fill_manual(values = colors(19))
p
ggsave(p, file="OUT/Plot/C_BCN_multi_stack_free_EXAE.jpeg", width=14.4, height=12)

#####?Ù»???????30??,???????????Ðµ?Ë³??#################
data_stage=subset(data_m, Process!="uNCE"&Process!="vNCE_EXAE"&Process!="aAE_lf"&Process!="bAE_er")
data_total=subset(data_m, Process=="vNCE_EXAE")
data_stage_neworder=data_stage
data_stage_neworder=data_stage_neworder[order(data_stage_neworder$NewID), ]
data_stage_neworder$Region=factor(data_stage_neworder$Region,levels = c(
 "Australia","Chile", "Mexico","Brazil","Indonesia",
 "Argentina","Malaysia","Philippines","Thailand","India",
 "United_States","United_Kingdom","Germany","Russia","China",
 "Canada","Japan","Finland","Spain","Korea",
 "Sweden","France","Italy","Austria","Poland",
 "New_Zealand","Norway","South_Africa","Portugal","Egypt"
  ))
data_total_neworder=data_total
data_total_neworder=data_total_neworder[order(data_total_neworder$NewID), ]
##??????????Ã»É¶?Ã£??????Ãµ???????Æ¥??
p=ggplot(data_stage_neworder,aes(Year,as.numeric(Value)/10^11,fill=Process))+
  geom_area(position="stack")+
  geom_line(data_total_neworder,mapping=aes(x=Year,y=as.numeric(Value)/10^11),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=14,face="bold"),
        axis.text.y=element_text(size=14,face="bold"),
        axis.title.x = element_text(size=14,face="bold"),
        axis.title.y = element_text(size=14,face="bold"))+
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  labs(text=element_text(family='calibri'),y="CO2 emission: Mt CO2-eq",x="Year")+
  facet_wrap(~NewID,ncol=5,scales="free")+
  scale_fill_manual(values = colors(19))+
  scale_y_continuous(labels = scales::number_format(accuracy = 0.01))
p
ggsave(p, file="OUT/Plot/Structure_stack_line_30NewOrder_EXAE231013-亿吨.jpg", dpi=600,width=14.4, height=12)








##########926#############???Åµ?SI?ï¡¿????+Ì¼???????????????Ê£?-?????????Å·?############################
##×¼??????
CarbonAll=read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "Carbon_BCN_PS")
CarbonAll=CarbonAll[1:1770,1:25]
data_m=melt(CarbonAll,id=c("NewID","Region","Year"),variable.name = "Process",value.name = "Value")#??Ô­Ê¼????????????Îª????????
data_m$Year=as.numeric(data_m$Year)
data_stage=subset(data_m, Process!="uNCE"&Process!="vNCE_EXAE")

colName = c( "#fd5f00", "#ebb481","#FF9933","#fecbaa","#feb64a",#Ì¼?æ´¢??,?????Å·?
             "#adbd37","#588133","#00a03e",#?É¼??×¶?
             "#fc9d9a",#??Ñ§Ä¾??????1
             "#FF717E",#??ÐµÄ¾??MP????2
             "#FF3B1D",#RP????3
             "#82052f",#??Ä¾??????4
             "#d8e9ef",#??×°Ö½PP??À¶1
             "#66c2cb",#PW,À¶2
             "#42B4E7" ,#??Ö½??À¶3
             "#6195C5",#????Ö½??À¶4
             "#005B9A",#????Ö½?Å£?À¶5
             "#8C489F",#PR????É«
             "#663300",#??Ë®
             "#C04D00",#????
             "white",#NCE
             "white"#NCE_EXAE
)#?????Æ£???????
colors = colorRampPalette(colName)#NÎª??Òª????É«??Á¿
image(x=1:22,y=1,z=as.matrix(1:22),col=colors(22))#Õ¹Ê¾Ò»??

#####????????30??,???????????Ðµ?Ë³??#################
data_NCE=subset(data_m, Process=="uNCE")
data_EXAE=subset(data_m, Process=="vNCE_EXAE")

p=ggplot(data_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(data_NCE,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="solid", col="red")+
  geom_line(data_EXAE,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=14,face="bold"),
        axis.text.y=element_text(size=14,face="bold"),
        axis.title.x = element_text(size=14,face="bold"),
        axis.title.y = element_text(size=14,face="bold"))+
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  labs(text=element_text(family='calibri'),y="CO2 emission: Mt CO2-eq",x="Year")+
  facet_wrap(~NewID,ncol=5,scales="free")+
  scale_fill_manual(values = colors(22))
p
ggsave(p, file="OUT/Plot/Structure_stack_line_30NewOrderSI0521.jpeg", width=14.4, height=12)






################623???Â´?????Ê±?Ã²????Ë£??????Î¿?######################
######?Ù»???Ò»??????############3
data1_stage=subset(data_stage, Region=="Brazil"|Region=="India"|
                     Region=="Indonesia"|Region=="China"|
                     Region=="Malaysia"|Region=="Korea"|
                     Region=="Chile"|Region=="Argentina"|
                     Region=="Thailand"|Region=="Philippines"|
                     Region=="Portugal"|Region=="Mexico"|
                     Region=="Egypt"|Region=="Russia")
data1_total=subset(data_total,Region=="Brazil"|Region=="India"|
                     Region=="Indonesia"|Region=="China"|
                     Region=="Malaysia"|Region=="Korea"|
                     Region=="Chile"|Region=="Argentina"|
                     Region=="Thailand"|Region=="Philippines"|
                     Region=="Portugal"|Region=="Mexico"|
                     Region=="Egypt"|Region=="Russia")
p=ggplot(data1_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(data1_total,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=12,face="bold"),
        axis.text.y=element_text(size=12,face="bold"),
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_fill_manual(values = colors(21))
p
ggsave(p, file="OUT/Plot/Structure_stack_line_type1.jpeg", width=16, height=7.5)

########################### ?Ù»??Ú¶??????? #######################################
data2_stage=subset(data_stage, Region=="United States"|Region=="New Zealand"|
                     Region=="Austria"|Region=="Italy"|
                     Region=="Poland"|Region=="Germany"|
                     Region=="South Africa"|Region=="Spain"|
                     Region=="Japan"|Region=="Australia")
data2_total=subset(data_total,Region=="United States"|Region=="New Zealand"|
                     Region=="Austria"|Region=="Italy"|
                     Region=="Poland"|Region=="Germany"|
                     Region=="South Africa"|Region=="Spain"|
                     Region=="Japan"|Region=="Australia")
p=ggplot(data2_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(data2_total,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=12,face="bold"),
        axis.text.y=element_text(size=12,face="bold"),
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_fill_manual(values = colors(21))
p
ggsave(p, file="OUT/Plot/Structure_stack_line_type2.jpeg", width=16, height=5)

########################### ?Ù»??????????? #######################################
data3_stage=subset(data_stage, Region=="Canada"|Region=="Finland"|
                     Region=="Sweden"|Region=="Norway"|
                     Region=="France"|Region=="United Kingdom")
data3_total=subset(data_total,Region=="Canada"|Region=="Finland"|
                     Region=="Sweden"|Region=="Norway"|
                     Region=="France"|Region=="United Kingdom")
p=ggplot(data3_stage,aes(Year,as.numeric(Value)/10^9,fill=Process))+
  geom_area(position="stack")+
  geom_line(data3_total,mapping=aes(x=Year,y=as.numeric(Value)/10^9),size=0.8,linetype="dashed")+
  theme(plot.margin=unit(rep(2,3,3,4),'cm'))+
  theme_bw() + theme(panel.grid = element_blank())+
  theme(text=element_text(family='calibri'),
        plot.title=element_text(size=14,hjust=0.5,face="bold"),
        axis.text.x=element_text(size=12,face="bold"),
        axis.text.y=element_text(size=12,face="bold"),
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
    ), # ????????x???????????????Í£?
    strip.text.y = element_text(
      size = 14, color = "red", face = "bold.italic"
    ) # ????????y???????????????Í£?
  )+
  facet_wrap(~Region,ncol=5,scales="free")+
  scale_fill_manual(values = colors(21))
p
ggsave(p, file="OUT/Plot/Structure_stack_line_type3.jpeg", width=16, height=5)
