##地图绘制

##Script Name: Paper Energy and Carbon
##腾出内存
options(java.parameters='-Xms8g')
gc()###内存不够时跑一下这个
memory.limit(1024000)

##引用要使用的包
library(RJDBC)
library(xlsx)
library(data.table)
library(ggplot2)
library(reshape2)      
library(do)
library(stringr)
library(cowplot)#图表拼接函数
library(RColorBrewer)
library(plyr)
library(mapdata)
library(ggmap)
library(maptools)
library(maps)
library(mapproj)
library(dplyr)

#############################准备工作###############################
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")##设置路径
#WMap=readShapePoly("INPUT/MapPlot/WorldMap/world map china line.shp")##这版地图没有小岛，还是用另一版比较放心
WMap=readShapePoly("INPUT/MapPlot/WorldMap/world countries.shp")
#NL=readShapeLines("INPUT/MapPlot/NineLines/NineLines.shp")##用不到了
#CP=readShapePoints("INPUT/MapPlot/Capitals/Capitals.shp")##用不到了

WMapData1=WMap@data##提取地图数据
WMapData2=data.frame(id=row.names(WMapData1),WMapData1)##给地图数据加个id方便后面连接绘图数据
#NLData1=NL@data
#NLData2=data.frame(id=row.names(NLData1),NLData1)
#CPData1=CP@data
#CPData2=data.frame(id=row.names(CPData1),CPData1)


##write.xlsx(WMapData2,file="OUT/MapData.xlsx",sheetName = "WMapData2",append=TRUE)##将上一条数据导出,为了给导入数据匹配上id
##在CarbonSheet和CarbonSheet_BioCarbonNeutral两个表格里的LcNCE_4Years工作表中加一列id，与上一行代码导出的数据相对应
##为了进一步在ggplot2包中绘图，需要把SpatialPolygonsDataFrame数据类型转化为真正的data.frame类型才可以。
##ggplot2包专门针对地理数据提供了特化版本的fortify函数来做这个工作
WMap1 = fortify(WMap)##转化为data.frame数据
WMapData = join(WMap1,WMapData2, type = "full") ##将地理信息与属性表连接起来
WMapData=WMapData[,-(128:129)]
WMapData=WMapData[,-(8:126)]
#NL1 = fortify(NL)##转化为data.frame数据
#NLData = join(NL1,NLData2, type = "full") ##将地理信息与属性表连接起来
#NLData=NLData[,-(128:129)]
#NLData=NLData[,-(8:126)]
#CP1 = fortify(CP)##转化为data.frame数据
#CPData = join(CP1,CPData2, type = "full") ##将地理信息与属性表连接起来
#CPData=CPData[,-(128:129)]
#CPData=CPData[,-(8:126)]




Item=c("1961","1980","2000","2019","1961-2019")

################################################## 4年批量出图 #################################################################
##包含生物质排放的批量出图——4年
LCNCE_4Years = read.xlsx("OUT/CarbonSheet.xlsx",sheetName = "LcNCE_4Years")
colnames(LCNCE_4Years)=c("id","Region","1961","1980","2000","2019","1961-2019")
LCNCE_4Years_L=melt(LCNCE_4Years,id=c("id","Region"),variable.name = "Year",value.name = "Value")
for(i in 5){
  LCNCE_Data=subset(LCNCE_4Years_L,Year==Item[i])
  LCNCE_Map = join(WMapData,LCNCE_Data, type = "full") ##将地理信息与属性表连接起来
  p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
    geom_polygon(aes(group=group,fill=Value),colour="#dadada",size=0.01)+
    scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions:\nBt CO2-eq")+ 
    #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
    theme (
      panel.grid=element_blank(),
      panel.background=element_blank(),
      axis.text=element_blank(),
      axis.ticks=element_blank(),
      axis.title=element_blank()
    )+
    ggtitle(Item[i])+
    theme(plot.title = element_text(size=15,color="#6b48ff",face = "bold",hjust = 0.5, vjust=0))+
    guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
  p
  
  ggsave(p,file=paste("OUT/Plot/RMap/LCNCE_",Item[i],".png",sep = "",collapse = NULL),width = 8,height = 4.5)
  
}

Item=c("1961","1980","2000","2019","1961-2019")
##不包含生物质排放的批量出图——4年
LCNCE_4Years = read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "LcNCE_4Years")
colnames(LCNCE_4Years)=c("id","Region","1961","1980","2000","2019","1961-2019")
LCNCE_4Years_L=melt(LCNCE_4Years,id=c("id","Region"),variable.name = "Year",value.name = "Value")
for(i in 5){
  LCNCE_Data=subset(LCNCE_4Years_L,Year==Item[i])
  LCNCE_Map = join(WMapData,LCNCE_Data, type = "full") ##将地理信息与属性表连接起来
  p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
    geom_polygon(aes(group=group,fill=Value),colour="#dadada",size=0.01)+
    scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions:\nBt CO2-eq")+ 
    #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
    theme (
      panel.grid=element_blank(),
      panel.background=element_blank(),
      axis.text=element_blank(),
      axis.ticks=element_blank(),
      axis.title=element_blank()
    )+
    ggtitle(Item[i])+
    theme(plot.title = element_text(size=15,color="#6b48ff",face = "bold",hjust = 0.5, vjust=0))+
    guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
  p
  
  ggsave(p,file=paste("OUT/Plot/RMap/LCNCE_BCN_",Item[i],".png",sep = "",collapse = NULL),width = 8,height = 4.5)
  
}


###############################################4个年份放在一张图里########################################################
##包含生物质的
LCNCE_4Years = read.xlsx("OUT/CarbonSheet.xlsx",sheetName = "LcNCE_4Years")
colnames(LCNCE_4Years)=c("id","Region","1961","1980","2000","2019","1961-2019")
LCNCE_4Years_L=melt(LCNCE_4Years,id=c("id","Region"),variable.name = "Year",value.name = "Value")
LCNCE_Map = join(WMapData,LCNCE_4Years_L, type = "full") ##将地理信息与属性表连接起来
LCNCE_Map=subset(LCNCE_Map,Year!="1961-2019"&Region!="MAX"&Region!="MIN")

p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
  geom_polygon(aes(group=group,fill=Value),colour="#dadada",size=0.01)+
  scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions:\nMt CO2-eq")+ 
  #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
  theme (
    panel.grid=element_blank(),
    panel.background=element_blank(),
    axis.text=element_blank(),
    axis.ticks=element_blank(),
    axis.title=element_blank()
  )+facet_wrap(~Year,ncol=2)+
  theme(
    strip.background = element_rect(color="white",fill="white"),
    strip.text = element_text(size=8,color="#003399",face = "bold"),
    legend.text = element_text(size = 8),
    legend.title = element_text(size=8)
  )+
  guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
p

ggsave(p,file="OUT/Plot/RMap/LCNCE_4Years.png",width=8,height = 4.5)

##不包含生物质的
LCNCE_4Years = read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "LcNCE_4Years_EXAE")
LCNCE_4Years=LCNCE_4Years[,1:7]
colnames(LCNCE_4Years)=c("id","Region","1961","1980","2000","2019","1961-2019")
LCNCE_4Years_L=melt(LCNCE_4Years,id=c("id","Region"),variable.name = "Year",value.name = "Value")
LCNCE_Map = join(WMapData,LCNCE_4Years_L, type = "full") ##将地理信息与属性表连接起来
LCNCE_Map=subset(LCNCE_Map,Year!="1961-2019"&Region!="MAX"&Region!="MIN")

p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
  geom_polygon(aes(group=group,fill=Value),colour="#808080",size=0.01)+
  scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),
                       na.value = "#EAEAEA",name="Carbon\nemissions:\nMt CO2-eq",
                       limits = c(-1, 200))+ 
  #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
  theme (
    panel.grid=element_blank(),
    panel.background=element_blank(),
    axis.text=element_blank(),
    axis.ticks=element_blank(),
    axis.title=element_blank()
  )+facet_wrap(~Year,ncol=2)+
  theme(
    strip.background = element_rect(color="white",fill="white"),
    strip.text = element_text(size=10,color="#003399",face = "bold"),
    legend.text = element_text(size = 8),
    legend.title = element_text(size=8)
  )+
  guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
p

ggsave(p,file="OUT/Plot/RMap/LcNCE_BCN_4Years_EXAE231106.pdf",width=16,height = 9)

#################################产量or消费量地图################
##读入数据
Flows=read.xlsx("OUT/Flows.xlsx",sheetName = "Flows")
Flows=Flows[,-1]
Flows=Flows[,-(3:4)]
##数据转换
Flows_L=melt(Flows,id=c("Region","Flow"),variable.name = "Year",value.name = "Value")
Flows_W=dcast(Flows_L,Region+Year~Flow,value.var = "Value")##转换成宽数据
Flows_4Years=subset(Flows_W,Year=="X1961"|Year=="X1980"|Year=="X2000"|Year=="X2019")
##纸张+纸浆产量
Flows_4Years$Value=Flows_4Years$F5+Flows_4Years$F7+Flows_4Years$F9+
  Flows_4Years$F15+Flows_4Years$F16+Flows_4Years$F17+Flows_4Years$F18+Flows_4Years$F19
##纸张产量
Flows_4Years$Value=Flows_4Years$F15+Flows_4Years$F16+Flows_4Years$F17+Flows_4Years$F18+Flows_4Years$F19
##纸张消费量
Flows_4Years$Value=Flows_4Years$F21+Flows_4Years$F22+Flows_4Years$F23+Flows_4Years$F24+Flows_4Years$F25

PP_Prd=Flows_4Years[,-(3:45)]
PP_Prd_W=dcast(PP_Prd,Region~Year,value.var = "Value")##转换成宽数据
##数据导出并添加id号码
PP_Prd_W=write.xlsx(PP_Prd_W,file = "OUT/Flows.xlsx",sheetName = "Paper_Con_W",append = TRUE)
##重新读入修改好的数据
PP_Prd_W=read.xlsx("OUT/Flows.xlsx",sheetName = "Paper_Con_W")
colnames(PP_Prd_W)=c("id","Region","1961","1980","2000","2019")
##重新整理用于画地图的数据
PP_Prd_L=melt(PP_Prd_W,id=c("id","Region"),variable.name = "Year",value.name = "Value")
PP_Prd_Map = join(WMapData,PP_Prd_L, type = "full") ##将地理信息与属性表连接起来
PP_Prd_Map=subset(PP_Prd_Map,Region!="MAX"&Region!="MIN")


##绘制地图
p=ggplot(PP_Prd_Map,aes(x=long,y=lat))+
  geom_polygon(aes(group=group,fill=Value/10^6),colour="#dadada",size=0.01)+
  scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions:\nMt CO2-eq")+ 
  #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
  theme (
    panel.grid=element_blank(),
    panel.background=element_blank(),
    axis.text=element_blank(),
    axis.ticks=element_blank(),
    axis.title=element_blank()
  )+facet_wrap(~Year,ncol=2)+
  theme(
    strip.background = element_rect(color="white",fill="white"),
    strip.text = element_text(size=10,color="#003399",face = "bold"),
    legend.text = element_text(size = 8),
    legend.title = element_text(size=8)
  )+
  guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
p

ggsave(p,file="OUT/Plot/RMap/Paper_Con_4Years.png",width=8,height = 4.5)



#####################################气泡地图#################################################
##四合一
LCNCE_4Years = read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "LcNCE_4Years")
colnames(LCNCE_4Years)=c("id","Region","1961","1980","2000","2019","1961-2019")
LCNCE_4Years_L=melt(LCNCE_4Years,id=c("id","Region"),variable.name = "Year",value.name = "Value")
LCNCE_Map = join(WMapData,LCNCE_4Years_L, type = "full") ##将地理信息与属性表连接起来
LCNCE_Map=subset(LCNCE_Map,Year!="1961-2019"&Region!="MAX"&Region!="MIN")

p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
  geom_polygon(aes(group=group,fill=Value),colour="#dadada",size=0.01)+
  geom_point(aes(x=size=Value, fill=Value, alpha=0.3), shape=21, colour="black")+
  scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions:\nMt CO2-eq")+ 
  #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
  theme (
    panel.grid=element_blank(),
    panel.background=element_blank(),
    axis.text=element_blank(),
    axis.ticks=element_blank(),
    axis.title=element_blank()
  )+facet_wrap(~Year,ncol=2)+
  theme(
    strip.background = element_rect(color="white",fill="white"),
    strip.text = element_text(size=10,color="#003399",face = "bold"),
    legend.text = element_text(size = 8),
    legend.title = element_text(size=8)
  )+
  guides(fill = guide_colourbar(barwidth = 0.5, barheight = 10))##调整颜色条大小
p
p1=p+ggplot(NL,aes(x=long,y=lat))+

ggsave(p,file="OUT/Plot/RMap/LCNCE_BCN_4Years.png",width=8,height = 4.5)






##########################59年批量出图#######################################################################
##不包含生物质排放的批量出图——多年
LCNCE_59Years = read.xlsx("OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName = "Carbon_ReCls")
LCNCE_59Years=LCNCE_59Years[,-(4:20)]
LCNCE_59Years=LCNCE_59Years[,-1]
LCNCE_59Years=subset(LCNCE_59Years,Year!="1961_2019")##去掉累积数据
LCNCE_59Years_W=dcast(LCNCE_59Years,Region~Year,value.var = "NCE")##转换成宽数据
LCNCE_59Years_W[31:32,1]=c("MAX","MIN")
LCNCE_59Years_W[31,2:60]=max(LCNCE_59Years_W[1:30,2:60])##加上最大值
LCNCE_59Years_W[32,2:60]=min(LCNCE_59Years_W[1:30,2:60])##加上最小值
LCNCE_59Years_W=cbind(LCNCE_4Years[1:32,1],LCNCE_59Years_W)#加上id
colnames(LCNCE_59Years_W)[1]="id"##修改列名
LCNCE_59Years_buding=matrix("NA",178,61)##新建一块补丁
LCNCE_59Years_buding=data.frame(LCNCE_59Years_buding)##将补丁转化成数据框
colnames(LCNCE_59Years_buding)=colnames(LCNCE_59Years_W)##给补丁重命名列名
LCNCE_59Years_W=rbind(LCNCE_59Years_W,LCNCE_59Years_buding)##将补丁加在原有数据上
LCNCE_59Years_W[33:210,1:2]=LCNCE_4Years[33:210,1:2]##将我们未研究的其他国家及id号码加进来
LCNCE_59Years_L=melt(LCNCE_59Years_W,id=c("id","Region"),variable.name = "Year",value.name = "Value")##转换成长数据，准备画图用
Item=colnames(LCNCE_59Years_W)[-(1:2)]##取出宽数据的列名作为循环变量
LCNCE_59Years_L_numeric=as.data.frame(lapply(LCNCE_59Years_L,as.numeric))
LCNCE_59Years_L$Value=LCNCE_59Years_L_numeric$Value

for(i in 1:59){
  LCNCE_Data=subset(LCNCE_59Years_L,Year==Item[i])
  LCNCE_Map = join(WMapData,LCNCE_Data, type = "full") ##将地理信息与属性表连接起来
  p=ggplot(LCNCE_Map,aes(x=long,y=lat))+
    geom_polygon(aes(group=group,fill=Value),colour="#dadada",size=0.01)+
    scale_fill_gradientn(colours=c("#005792","#FFF0BA","#fe7846","#f24a4a","#a10015"),na.value = "#EAEAEA",name="Carbon\nemissions")+ 
    #可以更改参数scale_fill_gradient2为两个颜色的渐变色,可以使用RGB颜色
    theme (
      panel.grid=element_blank(),
      panel.background=element_blank(),
      axis.text=element_blank(),
      axis.ticks=element_blank(),
      axis.title=element_blank()
    )+
    ggtitle(Item[i])+
    theme(plot.title = element_text(size=15,color="#6b48ff",face = "bold",hjust = 0.5, vjust=0))
  ggsave(p,file=paste("OUT/Plot/RMap/LCNCE_BCN_",Item[i],".png",sep = "",collapse = NULL))
  
}

