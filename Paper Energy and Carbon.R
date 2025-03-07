##Script Name: Paper Energy and Carbon
##腾出内存
options(java.parameters='-Xms8g')
gc()###内存不够时跑一下这个
memory.limit(1024000)
##导入使用到的包
install.packages("rJava") 
install.packages("xlsxjars") 
install.packages("xlsx") 
install.packages("RJDBC")
install.packages("plyr")
install.packages("openxlsx")
install.packages("xlsx")
install.packages("data.table")
install.packages("ggplot2")
install.packages("reshape2")
install.packages("do")
install.packages("stringr")
install.packages("cowplot")
install.packages("RColorBrewer")
install.packages("plyr")
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

################################################## 一、准备工作 ####################################################################
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")##设置路径
RG=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Regions")##读入国家名
Data_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "FAOSTAT")##将产量和进出口数据单独读入
Data=melt(Data_W,id=c("Region","Product","Item"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据

Flows_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName ='Flows')##读入物质流数据
Flows=melt(Flows_W,id=c("Region","Flow","Input","Output"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据

ELC_CI_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Elc_CI")##读入电力排放系数数据
ELC_CI=melt(ELC_CI_W,id=c("Region"),variable.name = "Year",value.name = "CI")#将原始宽型数据重组为长型数据
ELC_CI=subset(ELC_CI,Region!="World")##去掉没用的全球数据
write.xlsx(ELC_CI,file="OUT/Elc_CI.xlsx",sheetName="Elc_CI",append=TRUE)#导出数据

Product_P=subset(Data,(Product=="Mechanical wood pulp"|Product=="Chemical wood pulp"|
                         Product=="Pulp from fibres other than wood"|Product=="Recycled pulp"|
                         Product=="Newsprint"|Product=="Printing and writing papers"|
                         Product=="Household and sanitary papers"|Product=="Packaging paper and paperboard"|
                         Product=="Other paper and paperboard")&Item=="Production")##提取出产品产量数据，多个场景有需要

################################################## 二、进行各阶段的能源和碳排放核算 ##################################################

########################################### （一）【种植和采伐阶段】的碳存储与碳排放 ###############################################
########################################### 1. 【木材不可持续砍伐】引起的碳排放 ######################################################
##该部分在表格D:\坚果云同步文件夹\RBOOK\Global Paper\INPUT\DeforestData\Non-certified CarbonEmission.xlsx中计算
##之后复制进入CarbonEmission.xlsx中去，表格名称为CE_df_W，含义为carbon emission from deforestation宽数据
CE_df_W=read.xlsx("INPUT/ForestCarbon/ForestCarbon_Info.xlsx",sheetName = "CE_df")
CE_df_W=CE_df_W[(1:31),1:60]
names(CE_df_W)=c("Region",1961:2019)#命名列名
CE_df=melt(CE_df_W,id=c("Region"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
CE_df=CE_df[order(CE_df$Region),]##按照Region一列排序
write.xlsx(CE_df,file="OUT/CarbonEmission.xlsx",sheetName="CE_df",append=TRUE)#导出数据

########################################### 2. 【种植阶段】碳存储 #####################################################################
##说明：
##①该部分碳存储，经过生产制造使用，最终被固定在了产品存量或者填埋存量中
##②要扣除掉该阶段引发的碳排放，才算是净碳存储量
##本部分只是放在这里以防万一需要用到，目前在最终结果中并不体现
########################################### 2.1 浆木碳存储 ################################################################
PW=subset(Data,Region!="World"&Product=="Pulpwood"&Item=="Production")#木浆，浆木
CS_pw=PW[,1:4]
CS_pw$Value=PW$Value*0.26*44*1000/12#木材种植阶段碳汇，单位为kg。浆木单位为m3,0.26tonC/m3是含碳量，44/12是将C转化成CO2
CS_pw=CS_pw[,-3]
CS_pw=CS_pw[order(CS_pw$Region),]
write.xlsx(CS_pw,file="OUT/CarbonEmission.xlsx",sheetName="CS_pw",append=TRUE)#导出数据

########################################### 2.2 非木材碳存储 ################################################################
NW_CSI=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "NonWood_CSI")#各国一套数据
CS_nw=matrix(0,1770,3)
for (n in 1:30){
  NW_n=subset(Data,Region==RG[n,2]&Product=="Pulp from fibres other than wood"&Item=="Production")
  CS_nw[(59*n-58):(59*n),1]=RG[n,2]
  CS_nw[(59*n-58):(59*n),2]=1961:2019
  CS_nw[(59*n-58):(59*n),3]=NW_n$Value*NW_CSI$Value*1000#非木材生长碳汇，单位为kg
}
CS_nw=data.frame(CS_nw)#矩阵转化成数据框才能加列名
names(CS_nw)=c("Region","Year","Value")#添加列名
CS_nw=CS_nw[order(CS_nw$Region),]
write.xlsx(CS_nw,file="OUT/CarbonEmission.xlsx",sheetName="CS_nw",append=TRUE)#导出数据

########################################### （二）【原料收集阶段】能耗及碳排放 #######################
########################################### 1. 木材收集过程 ###########################################
########################################### 1.1 木材收集过程能耗 ###########################################
PW=subset(Data,Product=="Pulpwood"&Item=="Production")#木浆
E_wh=PW[,1:4]
E_wh$Value=PW$Value*0.54*(0.13+0.55)#木材收集过程能耗，单位GJ。0.54ton/m3是木材的密度，这里的估算使用所有木材的平均密度，不分区、不分树种。
E_wh=E_wh[,-3]
E_wh=E_wh[order(E_wh$Region),]
write.xlsx(E_wh,file="OUT/CarbonEmission.xlsx",sheetName="E_wh",append=TRUE)#导出数据

########################################### 1.2 木材收集过程碳排放 ###########################################

CE_wh=matrix(0,1770,3)
for (n in 1:30){
  PW_n=subset(Data,Region==RG[n,2]&Product=="Pulpwood")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  CE_wh[(59*n-58):(59*n),1]=RG[n,2]
  CE_wh[(59*n-58):(59*n),2]=1961:2019
  CE_wh[(59*n-58):(59*n),3]=PW_n$Value*0.54*0.13*74066.7/1000+PW_n$Value*0.54*0.55*ELC_CI_n$CI/1000#木材收集过程碳排放，单位kg
}
CE_wh=data.frame(CE_wh)#矩阵转化成数据框才能加列名
names(CE_wh)=c("Region","Year","Value")#添加列名
CE_wh=CE_wh[order(CE_wh$Region),]
write.xlsx(CE_wh,file="OUT/CarbonEmission.xlsx",sheetName="CE_wh",append=TRUE)#导出数据

########################################### 2. 非木材收集过程 ###########################################
########################################### 2.1 非木材收集过程能耗 ###########################################
NW_P=subset(Data,Product=="Pulp from fibres other than wood"&Item=="Production")##导入产量
E_nwh=NW_P[,1:4]
E_nwh$Value=(NW_P$Value)*0.3142#非木材原料收集能耗，单位GJ
E_nwh=E_nwh[,-3]
E_nwh=E_nwh[order(E_nwh$Region),]
write.xlsx(E_nwh,file="OUT/CarbonEmission.xlsx",sheetName="E_nwh",append=TRUE)#导出数据

########################################### 2.2 木材收集过程碳排放 ###########################################
CE_nwh=NW_P[,1:4]
CE_nwh$Value=(NW_P$Value)*23.3#非木材原料收集碳排，单位kg CO2。23.3kg CO2/t 非木材是提前算出来的单位非木材收集碳排放。
CE_nwh=CE_nwh[order(CE_nwh$Region), ]
CE_nwh=CE_nwh[,-3]
write.xlsx(CE_nwh,file="OUT/CarbonEmission.xlsx",sheetName="CE_nwh",append=TRUE)#导出数据

########################################### 3. 废纸收集过程 ###########################################
########################################### 3.1 废纸收集能耗 ###########################################
RP_P=subset(Data,Item=="Production"&Product=="Paper for recycling")
E_rpc=RP_P[,1:4]
E_rpc$Value=RP_P$Value*(0.0314+0.0702)#废纸收集能耗,GJ
E_rpc=E_rpc[,-3]
E_rpc=E_rpc[order(E_rpc$Region),]
write.xlsx(E_rpc,file="OUT/CarbonEmission.xlsx",sheetName="E_rpc",append=TRUE)#导出数据

########################################### 3.2 废纸收集碳排放 ###########################################
CE_rpc=matrix(0,1770,3)
for (n in 1:30){
  RP_n=subset(Data,Region==RG[n,2]&Product=="Paper for recycling"&Item=="Production")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  CE_rpc[(59*n-58):(59*n),1]=RG[n,2]
  CE_rpc[(59*n-58):(59*n),2]=1961:2019
  CE_rpc[(59*n-58):(59*n),3]=RP_n$Value*0.0314*74.0667+RP_n$Value*0.0702*ELC_CI_n$CI/1000#废纸收集过程碳排放，单位kg
}
CE_rpc=data.frame(CE_rpc)#矩阵转化成数据框才能加列名
names(CE_rpc)=c("Region","Year","Value")#添加列名
write.xlsx(CE_rpc,file="OUT/CarbonEmission.xlsx",sheetName="CE_rpc",append=TRUE)#导出数据

########################################### 4. 原料收集过程 ###########################################
########################################### 4.1 原料收集总能耗 ###########################################
E_he=E_wh[,1:4]
E_he$Value=E_wh$Value+E_nwh$Value+E_rpc$Value
write.xlsx(E_he,file="OUT/CarbonEmission.xlsx",sheetName="E_he",append=TRUE)#导出数据

########################################### 4.2 原料收集总碳排 ###########################################
CE_he=CE_wh[,1:2]
CE_he$Value=as.numeric(CE_wh$Value)+CE_nwh$Value+as.numeric(CE_rpc$Value)#原料收集总共的碳排放
write.xlsx(CE_he,file="OUT/CarbonEmission.xlsx",sheetName="CE_he",append=TRUE)#导出数据

########################################### （三）【化学品生产阶段】能耗及碳排放 ###########################################
########################################### 1. 化学品生产阶段能耗 ###########################################
Chem=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Chem")
Chem=subset(Chem,Region!="World")
E_chem=Chem[,1:4]
E_chem$Value=Product_P$Value*Chem$SEC#各种纸浆和纸张药品投入的能源消耗，GJ
E_chem=E_chem[,-3]
write.xlsx(E_chem,file="OUT/CarbonEmission.xlsx",sheetName="E_chem",append=TRUE)#导出数据

########################################### 2. 产品合并，计算总量 ###########################################
E_chem_W=dcast(E_chem,Region+Year~Product,value.var = "Value")
E_chem_W$SUM=rowSums(E_chem_W[,3:11])
E_chem_P=E_chem_W[,-(3:11)]
write.xlsx(E_chem_P,file="OUT/CarbonEmission.xlsx",sheetName="E_chem_P",append=TRUE)#导出数据

########################################### 3. 碳排放 ###########################################
CE_chem=Chem[,1:4]
CE_chem$Value=Product_P$Value*Chem$CI#各种纸浆和纸张药品投入的碳排放,kg
CE_chem=CE_chem[,-3]
write.xlsx(CE_chem,file="OUT/CarbonEmission.xlsx",sheetName="CE_chem",append=TRUE)#导出数据12

########################################### （四）废水处理过程碳排放 ###########################################
Product_P=Product_P[order(Product_P$Region), ]
Product_P=Product_P[order(Product_P$Year), ]
COD_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "WtTrt_COD")
COD=melt(COD_W,id=c("Region","Year"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
COD=subset(COD,Region!="World")
COD=COD[order(COD$Region), ]
COD=COD[order(COD$Year), ]
Removal_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "WtTrt_Removal")
Removal=melt(Removal_W,id=c("Region","Year"),variable.name = "Product",value.name = "Value")
Removal=subset(Removal,Region!="World")
Removal=Removal[order(Removal$Region), ]
Removal=Removal[order(Removal$Year), ]
MCF_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "WtTrt_MCF")
MCF=melt(MCF_W,id=c("Region","Year"),variable.name = "Product",value.name = "Value")
MCF=subset(MCF,Region!="World")
MCF=MCF[order(MCF$Region), ]
MCF=MCF[order(MCF$Year), ]
########################################### 1. 排放到空气中的甲烷 ###########################################
########################################### 1.1 排放到空气中的总量 ###########################################
CE_wttrt=COD[,1:3]##排放到空气的碳
CE_wttrt$Value=Product_P$Value*COD$Value*Removal$Value*MCF$Value*0.25*25#乘以25转化为CO2当量，单位是kg
write.xlsx(CE_wttrt,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt",append=TRUE)#导出数据15

########################################### 1.2 制浆过程排放到空气中的甲烷（kg CO2-eq） ###########################################
CE_wttrt_Pulp=subset(CE_wttrt,(Product=="Mechanical.wood.pulp"|Product=="Chemical.wood.pulp"|
                                 Product=="Pulp.from.fibres.other.than.wood"|Product=="Recovered.fibre.pulp"))#导入纸张净进口数据
CE_wttrt_Pulp_W=dcast(CE_wttrt_Pulp,Region+Year~Product,value.var = "Value")
CE_wttrt_Pulp_W$Value=rowSums(CE_wttrt_Pulp_W[,3:6])
CE_wttrt_Pulp_SUM=CE_wttrt_Pulp_W[,-(3:6)]
write.xlsx(CE_wttrt_Pulp_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt_Pulp",append=TRUE)#导出数据

########################################### 1.3 造纸过程排放到空气中的甲烷 ###########################################
CE_wttrt_Paper=subset(CE_wttrt,(Product=="Newsprint"|Product=="Printing.and.writing.papers"|
                                  Product=="Household.and.sanitary.papers"|Product=="Packaging.paper.and.paperboard"|
                                  Product=="Other.paper.and.paperboard"))#导入数据
CE_wttrt_Paper_W=dcast(CE_wttrt_Paper,Region+Year~Product,value.var = "Value")
CE_wttrt_Paper_W$Value=rowSums(CE_wttrt_Paper_W[,3:7])
CE_wttrt_Paper_SUM=CE_wttrt_Paper_W[,-(3:7)]
write.xlsx(CE_wttrt_Paper_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt_Paper",append=TRUE)#导出数据15

########################################### 2. 排放到淤泥或水体中的碳(kg CO2-eq) ###########################################
########################################### 2.1 排放到淤泥或水体中的碳总量(kg CO2-eq) ###########################################
CE_wttrt_ToNature=COD[,1:3]
CE_wttrt_ToNature$Value=((Product_P$Value*COD$Value-19.385)*(1/1.8646)-CE_wttrt$Value*12/(25*16))*44/12#单位是kg CO2
CE_wttrt_ToNature[CE_wttrt_ToNature<0]=0##该替换负值为零的方法在这个地方失灵了
##所以只能用以下条件判断语句实现该功能了
for (n in 1:1770){
  if(CE_wttrt_ToNature$Value[n]<0){
    CE_wttrt_ToNature$Value[n]=0
  }else{
    CE_wttrt_ToNature$Value[n]=CE_wttrt_ToNature$Value[n]
  }
}##成功！

write.xlsx(CE_wttrt_ToNature,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt_ToNature",append=TRUE)#导出数据15

########################################### 2.2 制浆过程排放到淤泥或水体中的碳(kg CO2-eq) ###########################################
CE_wttrt_ToNature_Pulp=subset(CE_wttrt_ToNature,(Product=="Mechanical.wood.pulp"|Product=="Chemical.wood.pulp"|
                                 Product=="Pulp.from.fibres.other.than.wood"|Product=="Recovered.fibre.pulp"))#导入纸张净进口数据
CE_wttrt_ToNature_Pulp_W=dcast(CE_wttrt_ToNature_Pulp,Region+Year~Product,value.var = "Value")
CE_wttrt_ToNature_Pulp_W$Value=rowSums(CE_wttrt_ToNature_Pulp_W[,3:6])
CE_wttrt_ToNature_Pulp_SUM=CE_wttrt_ToNature_Pulp_W[,-(3:6)]
write.xlsx(CE_wttrt_ToNature_Pulp_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt_ToNature_Pulp",append=TRUE)#导出数据15

########################################### 2.3 造纸过程排放到淤泥或水体中的碳(kg CO2-eq) ###########################################
CE_wttrt_ToNature_Paper=subset(CE_wttrt_ToNature,(Product=="Newsprint"|Product=="Printing.and.writing.papers"|
                                  Product=="Household.and.sanitary.papers"|Product=="Packaging.paper.and.paperboard"|
                                  Product=="Other.paper.and.paperboard"))#导入纸张净进口数据
CE_wttrt_ToNature_Paper_W=dcast(CE_wttrt_ToNature_Paper,Region+Year~Product,value.var = "Value")
CE_wttrt_ToNature_Paper_W$Value=rowSums(CE_wttrt_ToNature_Paper_W[,3:7])
CE_wttrt_ToNature_Paper_SUM=CE_wttrt_ToNature_Paper_W[,-(3:7)]
write.xlsx(CE_wttrt_ToNature_Paper_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CE_wttrt_ToNature_Paper",append=TRUE)#导出数据15

########################### （五）使用和回收过程的碳汇和碳排放 #########################
########################## 0. 扣除非生物质部分的流量 #########################
F14=subset(Flows,Flow=="F14")
F14=F14[order(F14$Region), ]
F15=subset(Flows,Flow=="F15")
F15=F15[order(F15$Region), ]
F16=subset(Flows,Flow=="F16")
F16=F16[order(F16$Region), ]
F17=subset(Flows,Flow=="F17")
F17=F17[order(F17$Region), ]
F18=subset(Flows,Flow=="F18")
F18=F18[order(F18$Region), ]
F19=subset(Flows,Flow=="F19")
F19=F19[order(F19$Region), ]
F21=subset(Flows,Flow=="F21")
F21=F21[order(F21$Region), ]
F22=subset(Flows,Flow=="F22")
F22=F22[order(F22$Region), ]
F23=subset(Flows,Flow=="F23")
F23=F23[order(F23$Region), ]
F24=subset(Flows,Flow=="F24")
F24=F24[order(F24$Region), ]
F25=subset(Flows,Flow=="F25")
F25=F25[order(F25$Region), ]
##导入5种纸产品的净进口量
F38=subset(Flows,Flow=="F38")##NP
F38=F38[order(F38$Region), ]
F39=subset(Flows,Flow=="F39")##PW
F39=F39[order(F39$Region), ]
F40=subset(Flows,Flow=="F40")##HS
F40=F40[order(F40$Region), ]
F41=subset(Flows,Flow=="F41")##PP
F41=F41[order(F41$Region), ]
F42=subset(Flows,Flow=="F42")##OP
F42=F42[order(F42$Region), ]

MBP_AM=read.xlsx("INPUT/PaperMFA/PaperMFA_China.xlsx",sheetName = "MBP_AM")#MBP_AM是物质平衡的输入参数,分配参数
AM=MBP_AM[1:20,3:61]#AM是纸浆配比数据,这里，我们仅使用第八列，即书写用纸中的非生物质比例

##五种纸产品中的生物质部分
F21_Bio=F15[,1:5]
F21_Bio$Flow="F21_Bio"
F21_Bio$Value=(F15$Value+F38$Value)*0.9
F22_Bio=F16[,1:5]
F22_Bio$Flow="F22_Bio"
F22_Bio$Value=(F16$Value+F39$Value)*(1-as.numeric(AM[8,]))
F23_Bio=F17[,1:5]
F23_Bio$Flow="F23_Bio"
F23_Bio$Value=(F17$Value+F40$Value)*1
F24_Bio=F18[,1:5]
F24_Bio$Flow="F24_Bio"
F24_Bio$Value=(F18$Value+F41$Value)*0.9
F25_Bio=F19[,1:5]
F25_Bio$Flow="F25_Bio"
F25_Bio$Value=(F19$Value+F42$Value)*0.9
  

TPC_Bio=F14[,1:5]
TPC_Bio$Flow="TPC_Bio"
TPC_Bio$Input="Papermaking"
TPC_Bio$Output="Paper and paperboard"
TPC_Bio$Value=F21_Bio$Value+F22_Bio$Value+F23_Bio$Value+F24_Bio$Value+F25_Bio$Value

##废纸消费量（即废纸制浆原料投入量）中的生物质
F4=subset(Flows,Flow=="F4")
F4=F4[order(F4$Region), ]
F4_Bio=F4[,1:5]
F4_Bio$Flow="F4_Bio"
F4_Bio$Value=F4$Value*(1-(F21$Value*0.1+F22$Value*as.numeric(AM[8,])+F23$Value*0+F24$Value*0.1+F25$Value*0.1)/
                          (F21$Value+F22$Value+F23$Value+F24$Value+F25$Value))

F43=subset(Flows,Flow=="F43")
F43=F43[order(F43$Region), ]

##废纸产生量中的生物质
PFR=F4##PFR指的是PaPer for Recycling
PFR$Value=F4$Value-F43$Value##减去净进口得到产生量，等于F27
PFR_Bio=PFR
PFR_Bio$Flow="PFR_Bio"
PFR_Bio$Value=PFR$Value*(1-(F21$Value*0.1+F22$Value*as.numeric(AM[8,])+F23$Value*0+F24$Value*0.1+F25$Value*0.1)/
                           (F21$Value+F22$Value+F23$Value+F24$Value+F25$Value))##假设其中的非生物质比例与当年消费的纸产品中的非生物质比例的均值一致


F9=subset(Flows,Flow=="F9")
F9=F9[order(F9$Region), ]
F10=subset(Flows,Flow=="F10")
F10=F10[order(F10$Region), ]
F11=subset(Flows,Flow=="F11")
F11=F11[order(F11$Region), ]
F12=subset(Flows,Flow=="F12")
F12=F12[order(F12$Region), ]
F13=subset(Flows,Flow=="F13")
F13=F13[order(F13$Region), ]
F26=subset(Flows,Flow=="F26")
F26=F26[order(F26$Region), ]
F27=subset(Flows,Flow=="F27")
F27=F27[order(F27$Region), ]
F28=subset(Flows,Flow=="F28")
F28=F28[order(F28$Region), ]
F29=subset(Flows,Flow=="F29")
F29=F29[order(F29$Region), ]
F30=subset(Flows,Flow=="F30")
F30=F30[order(F30$Region), ]
F31=subset(Flows,Flow=="F31")
F31=F31[order(F31$Region), ]


F15_Bio=F15[,1:5]
F15_Bio$Flow="F15_Bio"
F15_Bio$Value=F15$Value*0.9
F16_Bio=F16[,1:5]
F16_Bio$Flow="F16_Bio"
F16_Bio$Value=F16$Value*(1-as.numeric(AM[8,]))
F17_Bio=F17[,1:5]
F17_Bio$Flow="F17_Bio"
F17_Bio$Value=F17$Value*1
F18_Bio=F18[,1:5]
F18_Bio$Flow="F18_Bio"
F18_Bio$Value=F18$Value*0.9
F19_Bio=F19[,1:5]
F19_Bio$Flow="F19_Bio"
F19_Bio$Value=F19$Value*0.9

F20=subset(Flows,Flow=="F20")
F20=F20[order(F20$Region), ]
F20_Bio=F20[,1:5]
F20_Bio$Flow="F20_Bio"
F20_Bio$Value=(F11$Value+F12$Value+F13$Value)-CE_wttrt_Paper_SUM$Value*12/(25*16*0.45*1000)-
  CE_wttrt_ToNature_Paper_SUM$Value*12/(44*0.45*1000)-
  (F15_Bio$Value+F16_Bio$Value+F17_Bio$Value+F18_Bio$Value+F19_Bio$Value)
  #Papermaking to Paper for recycling

F26_Bio=F26[,1:5]
F26_Bio$Flow="F26_Bio"
F26_Bio$Value=0
F27_Bio=F27[,1:5]
F27_Bio$Flow="F27_Bio"
F27_Bio$Value=0
F28_Bio=F28[,1:5]
F28_Bio$Flow="F28_Bio"
F28_Bio$Value=0
F29_Bio=F29[,1:5]
F29_Bio$Flow="F29_Bio"
F29_Bio$Value=0
F30_Bio=F30[,1:5]
F30_Bio$Flow="F30_Bio"
F30_Bio$Value=0
F31_Bio=F31[,1:5]
F31_Bio$Flow="F31_Bio"
F31_Bio$Value=0


for (n in 1:30){
  MBP_YR=read.xlsx(paste("INPUT/PaperMFA/PaperMFA_",RG[n,2],".xlsx",sep = "",collapse = NULL),sheetName = "MBP_YR")
  YR=MBP_YR[1:10,2:60]#YR是产出率表格的参数数据
  F26_Bio$Value[(59*n-58):(59*n)]=TPC_Bio$Value[(59*n-58):(59*n)]*0.09#Use to Stock
  F27_Bio$Value[(59*n-58):(59*n)]=PFR_Bio$Value[(59*n-58):(59*n)]#Use to Paper for recycling
  F28_Bio$Value[(59*n-58):(59*n)]=(TPC_Bio$Value[(59*n-58):(59*n)]+F20_Bio$Value[(59*n-58):(59*n)]-F26_Bio$Value[(59*n-58):(59*n)]-F27_Bio$Value[(59*n-58):(59*n)])*as.numeric(YR[8,])#Use to Incineration
  F29_Bio$Value[(59*n-58):(59*n)]=(TPC_Bio$Value[(59*n-58):(59*n)]+F20_Bio$Value[(59*n-58):(59*n)]-F26_Bio$Value[(59*n-58):(59*n)]-F27_Bio$Value[(59*n-58):(59*n)])*as.numeric(YR[6,])#Use to Energy recovery
  F30_Bio$Value[(59*n-58):(59*n)]=(TPC_Bio$Value[(59*n-58):(59*n)]+F20_Bio$Value[(59*n-58):(59*n)]-F26_Bio$Value[(59*n-58):(59*n)]-F27_Bio$Value[(59*n-58):(59*n)])*as.numeric(YR[7,])#Use to Non-energy recovery
  F31_Bio$Value[(59*n-58):(59*n)]=TPC_Bio$Value[(59*n-58):(59*n)]+F20_Bio$Value[(59*n-58):(59*n)]-F26_Bio$Value[(59*n-58):(59*n)]-F27_Bio$Value[(59*n-58):(59*n)]-
    F28_Bio$Value[(59*n-58):(59*n)]-F29_Bio$Value[(59*n-58):(59*n)]-F30_Bio$Value[(59*n-58):(59*n)]#Use to Landfill
}

Flows_Bio=rbind(F4_Bio,F15_Bio,F16_Bio,F17_Bio,F18_Bio,F19_Bio,F20_Bio,F21_Bio,F22_Bio,F23_Bio,F24_Bio,F25_Bio,F26_Bio,
               F27_Bio,F28_Bio,F29_Bio,F30_Bio,F31_Bio)
Flows_Bio_W=dcast(Flows_Bio,Region+Flow+Input+Output~Year,value.var = "Value")
Flows_Bio_W[Flows_Bio_W<0]=0
Flows_Bio_W[is.na(Flows_Bio_W)]=0
write.xlsx(Flows_Bio_W,file="OUT/Flows.xlsx",sheetName="Flows_Bio",append=TRUE)#导出数据

########################### 1. 产品消费量 #########################
########################### 1.1 纸产品消费量 #########################
Product_NI=subset(Data,(Product=="Mechanical wood pulp"|Product=="Chemical wood pulp"|
                          Product=="Pulp from fibres other than wood"|Product=="Recycled pulp"|
                          Product=="Newsprint"|Product=="Printing and writing papers"|
                          Product=="Household and sanitary papers"|Product=="Packaging paper and paperboard"|
                          Product=="Other paper and paperboard")&Item=="Net import"&Region!="World")#导入纸张净进口数据
Product_C=Product_P[,1:4]
Product_C$Value=Product_P$Value+Product_NI$Value#消费量=生产量+净进口量
Product_C$Item="Consumption"
Paper_C=subset(Product_C,Product=="Newsprint"|Product=="Printing and writing papers"|
                 Product=="Household and sanitary papers"|Product=="Packaging paper and paperboard"|
                 Product=="Other paper and paperboard")
Paper_C=Paper_C[,-3]
Paper_C_W=dcast(Paper_C,Region+Year~Product,value.var = "Value")
write.xlsx(Paper_C_W,file="OUT/Product_Data.xlsx",sheetName="Paper_C_W",append=TRUE)#导出数据

########################## 1.2 纸浆产品消费量 ###########################################
Pulp_C=subset(Product_C,Product=="Mechanical wood pulp"|Product=="Chemical wood pulp"|
                Product=="Pulp from fibres other than wood"|Product=="Recycled pulp")
Pulp_C=Pulp_C[,-3]
Pulp_C_W=dcast(Pulp_C,Region+Year~Product,value.var = "Value")
write.xlsx(Pulp_C_W,file="OUT/Product_Data.xlsx",sheetName="Pulp_C_W",append=TRUE)#导出数据16

######################### 1.3 消费品形成产品存量——碳汇（选用CS_ps_Bio_SUM）#############
##产品存量碳汇分产品计算
##排除了非生物质部分
Paper_C_Bio=rbind(F21_Bio,F22_Bio,F23_Bio,F24_Bio,F25_Bio)
Paper_C_Bio=Paper_C_Bio[,-(2:3)]
names(Paper_C_Bio)=c("Region","Product","Year","Value")
Paper_C_Bio$Year=1961:2019
CS_ps_Bio=Paper_C_Bio[,1:3]
CS_ps_Bio$Value=Paper_C_Bio$Value*0.09*0.45*44*1000/12#单位为kg CO2
write.xlsx(CS_ps_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CS_ps_Bio",append=TRUE)#导出数据
##产品存量碳汇总量计算
#####排除了非生物质########
CS_ps_Bio_W=dcast(CS_ps_Bio,Region+Year~Product,value.var = "Value")
CS_ps_Bio_SUM=CS_ps_Bio_W[,1:2]
CS_ps_Bio_SUM$Value=rowSums(CS_ps_Bio_W[,3:7])
write.xlsx(CS_ps_Bio_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CS_ps_Bio_SUM",append=TRUE)#导出数据


########################## 2.生活废物回收库存碳汇（选用CS_rc_Bio）#########################
#相关计算方法没有完全确定,不过量不大，不太影响结论
########排除了非生物质###########
CS_rc_Bio=matrix(0,1770,3)
CS_rc_Bio_n=matrix(0,59,1)
for (n in 1:30){
  Recycling_P=F4_Bio$Value[(59*n-58):(59*n)]
  for (t in 2:59) {
    CS_rc_Bio_n[t,1]=(Recycling_P[t]-Recycling_P[t-1])*0.45*1000*44/12#单位为kg CO2
  }
  CS_rc_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CS_rc_Bio[(59*n-58):(59*n),2]=1961:2019
  CS_rc_Bio[(59*n-58):(59*n),3]=CS_rc_Bio_n
}
CS_rc_Bio=data.frame(CS_rc_Bio)#矩阵转化成数据框才能加列名
names(CS_rc_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CS_rc_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CS_rc_Bio",append=TRUE)#导出数据

########################## 3.生活废物燃烧产生的碳排放（选用CE_inc_Bio）#########################
################排除了非生物质###########
CE_inc_Bio=matrix(0,1770,3)
for (n in 1:30){
  CE_inc_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CE_inc_Bio[(59*n-58):(59*n),2]=1961:2019
  CE_inc_Bio[(59*n-58):(59*n),3]=F28_Bio$Value[(59*n-58):(59*n)]*0.45*1000*44/12#最终单位为kg.
}
CE_inc_Bio=data.frame(CE_inc_Bio)#矩阵转化成数据框才能加列名
names(CE_inc_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CE_inc_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CE_inc_Bio",append=TRUE)#导出数据

##########################4.生活废物能源回收###################
################### 4.1 生活废物能源回收部分的碳排放（选用CE_er_mw_Bio）#########################
###########排除了非生物质############
CE_er_mw_Bio=matrix(0,1770,3)
for (n in 1:30){
  CE_er_mw_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CE_er_mw_Bio[(59*n-58):(59*n),2]=1961:2019
  CE_er_mw_Bio[(59*n-58):(59*n),3]=F29_Bio$Value[(59*n-58):(59*n)]*0.45*1000*44/12#最终单位为kg
}
CE_er_mw_Bio=data.frame(CE_er_mw_Bio)#矩阵转化成数据框才能加列名
names(CE_er_mw_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CE_er_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CE_er_mw_Bio",append=TRUE)#导出数据

########################## 4.2 生活废物回收的能源（选用E_er_mw）#########################
E_er_mw=matrix(0,1770,3)
for (n in 1:30){
  F29=subset(Flows,Region==RG[n,2]&Flow=="F29")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  E_er_mw[(59*n-58):(59*n),1]=RG[n,2]
  E_er_mw[(59*n-58):(59*n),2]=1961:2019
  F29[is.na(F29)]=0
  E_er_mw[(59*n-58):(59*n),3]=F29$Value*13#最终单位GJ
}
E_er_mw=data.frame(E_er_mw)#矩阵转化成数据框才能加列名
names(E_er_mw)=c("Region","Year","Value")#添加列名
write.xlsx(E_er_mw,file="OUT/CarbonEmission.xlsx",sheetName="E_er_mw",append=TRUE)


########################## 4.3 能源回收避免的碳排放——生活废物（选用AE_er_mw）#########################
AE_er_mw=matrix(0,1829,3)
for (n in 1:30){
  F29=subset(Flows,Region==RG[n,2]&Flow=="F29")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  AE_er_mw[(59*n-58):(59*n),1]=RG[n,2]
  AE_er_mw[(59*n-58):(59*n),2]=1961:2019
  AE_er_mw[(59*n-58):(59*n),3]=F29$Value*0.013*0.25*ELC_CI_n$CI#最终单位为kg CO2
}
AE_er_mw=data.frame(AE_er_mw)#矩阵转化成数据框才能加列名
names(AE_er_mw)=c("Region","Year","Value")#添加列名
write.xlsx(AE_er_mw,file="OUT/CarbonEmission.xlsx",sheetName="AE_er_mw",append=TRUE)#导出数据

##########################5.生活废物填埋#########################
######################5.0 生活废物填埋释放CO2（选用CO2_lf_mw_Bio）#############################################
CO2_lf_mw_Bio=matrix(0,1770,3)
CO2_lf_mw_Bio_n=matrix(0,59,1)
for (n in 1:30){
  Landfill_P=F31_Bio$Value[(59*n-58):(59*n)]
  t=seq(1961,2019,1)
  for (i in 2:59) {
    j=1:i
    CO2_lf_mw_Bio_n[i,1]=sum((1-exp(-0.05))*Landfill_P[j]*0.7*0.45*0.5*0.5*
                              (44/12)*exp(-0.05*(t[i]-t[j])))*1000#单位为kg
  }
  CO2_lf_mw_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CO2_lf_mw_Bio[(59*n-58):(59*n),2]=1961:2019
  CO2_lf_mw_Bio[(59*n-58):(59*n),3]=CO2_lf_mw_Bio_n
}
CO2_lf_mw_Bio=data.frame(CO2_lf_mw_Bio)#矩阵转化成数据框才能加列名
names(CO2_lf_mw_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CO2_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CO2_lf_mw_Bio",append=TRUE)#导出数据

###################### 5.1 生活废物填埋释放CH4（选用CE_lf_mw_Bio）#############################################
###########排除了非生物质############
CE_lf_mw_Bio=matrix(0,1770,3)
CE_lf_mw_Bio_n=matrix(0,59,1)
for (n in 1:30){
  Landfill_P=F31_Bio$Value[(59*n-58):(59*n)]
  t=seq(1961,2019,1)
  for (i in 2:59) {
    j=1:i
    CE_lf_mw_Bio_n[i,1]=sum((1-exp(-0.05))*Landfill_P[j]*0.7*0.45*0.5*0.5*
                              (16/12)*exp(-0.05*(t[i]-t[j]))*
                              (1-0.25)*(1-0.05))*25*1000#单位为kg
  }
  CE_lf_mw_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CE_lf_mw_Bio[(59*n-58):(59*n),2]=1961:2019
  CE_lf_mw_Bio[(59*n-58):(59*n),3]=CE_lf_mw_Bio_n
}
CE_lf_mw_Bio=data.frame(CE_lf_mw_Bio)#矩阵转化成数据框才能加列名
names(CE_lf_mw_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CE_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CE_lf_mw_Bio",append=TRUE)#导出数据

########################### 5.2 生活废物填埋产生CH4的收集 #########################
########################### 5.2.1 生活废物填埋产生CH4收集后燃烧碳排放（选用CECP_lf_mw_Bio）#########################
########排除了非生物质###########
CECP_lf_mw_Bio=CE_lf_mw_Bio[,1:2]
CECP_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio$Value)*0.25*44/(25*(1-0.25)*(1-0.05)*16)#最终单位为kg CH4
write.xlsx(CECP_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CECP_lf_mw_Bio",append=TRUE)#导出数据

########################## 5.2.2 生活废物填埋CH4收集供应的能量：GJ（选用E_lf_mw_Bio）#########################
#####排除了非生物质########
E_lf_mw_Bio=CECP_lf_mw_Bio[1:2]
E_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio$Value)*0.25/(25*(1-0.25)*(1-0.05))*0.048*1000*0.35#产生的甲烷单位为kg, 天然气热值为0.048TJ/ton，假设该部分能量要转换成最后结果单位为GJ
write.xlsx(E_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="E_lf_mw_Bio",append=TRUE)#导出数据

#######################5.2.3生活废物填埋产生的CH4收集量避免的排放（选用AE_lf_mw_Bio）#########################
########排除了非生物质###########
AE_lf_mw_Bio=CE_lf_mw_Bio[1:2]
ELC_CI=subset(ELC_CI,Region!="World")
AE_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio$Value)*0.25/(25*(1-0.25)*(1-0.05))*0.048*0.35*ELC_CI$CI/1000#CCP_lf单位为kg, 天然气热值为0.048TJ/ton，发电效率0.35，ELC_CI单位为kg CO2/TJ，最后结果单位为kg
write.xlsx(AE_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="AE_lf_mw_Bio",append=TRUE)#导出数据


##########################5.3 生活废物填埋逸出的CO2（CESC_lf_mw_Bio）#####################################################
########排除了非生物质##########
CESC_lf_mw_Bio=CE_lf_mw_Bio[,1:2]
CESC_lf_mw_Bio$Value=(as.numeric(CE_lf_mw_Bio$Value)/25)*(0.05/(1-0.05))*44/16#最终单位为kg
write.xlsx(CESC_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CESC_lf_mw_Bio",append=TRUE)#导出数据28

################5.4 碳汇——（选用CS_lf_mw_Bio）######
########使用每年的填埋处理总量（inflow）减掉释放出去的部分（Outflow）得到当年的填埋存量净增量############
##假设1961年之前的存量为0
########排除了非生物质###########
CS_lf_mw_Bio=matrix(0,1770,3)
for (n in 1:30){
  CS_lf_mw_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CS_lf_mw_Bio[(59*n-58):(59*n),2]=1961:2019
  F31_Bio[is.na(F31_Bio)]=0
  CS_lf_mw_Bio[(59*n-58):(59*n),3]=(F31_Bio$Value[(59*n-58):(59*n)]*0.45*1000-as.numeric(CO2_lf_mw_Bio$Value[(59*n-58):(59*n)])*12/44-
    as.numeric(CE_lf_mw_Bio$Value[(59*n-58):(59*n)])*12/(25*16)-CECP_lf_mw_Bio$Value[(59*n-58):(59*n)]*12/44-
    CESC_lf_mw_Bio$Value[(59*n-58):(59*n)]*12/44)*44/12#最终单位为kg
}
CS_lf_mw_Bio=data.frame(CS_lf_mw_Bio)#矩阵转化成数据框才能加列名
names(CS_lf_mw_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CS_lf_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CS_lf_mw_Bio",append=TRUE)#导出数据

############ 6. 生活废物非能源回收所含碳################
##排除了生物质
CS_ner_mw_Bio=matrix(0,1770,3)
for (n in 1:30){
  CS_ner_mw_Bio[(59*n-58):(59*n),1]=RG[n,2]
  CS_ner_mw_Bio[(59*n-58):(59*n),2]=1961:2019
  F30_Bio[is.na(F30_Bio)]=0
  CS_ner_mw_Bio[(59*n-58):(59*n),3]=F30_Bio$Value[(59*n-58):(59*n)]*0.45*1000*44/12#最终单位为kg CO2
}
CS_ner_mw_Bio=data.frame(CS_ner_mw_Bio)#矩阵转化成数据框才能加列名
names(CS_ner_mw_Bio)=c("Region","Year","Value")#添加列名
write.xlsx(CS_ner_mw_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CS_ner_mw_Bio",append=TRUE)#导出数据


###################### （六）工厂废物去向 ###################
########################## 1.工业废物能源回收#########################
########################## 1.1工业废物能源回收部分的碳排放（CE_er_iw）#########################
##本来就不包含非生物质
CE_er_iw=matrix(0,1770,3)
for (n in 1:30){
  F32=subset(Flows,Region==RG[n,2]&Flow=="F32")
  CE_er_iw[(59*n-58):(59*n),1]=RG[n,2]
  CE_er_iw[(59*n-58):(59*n),2]=1961:2019
  CE_er_iw[(59*n-58):(59*n),3]=F32$Value*1000*0.45*44/12#最终单位为kgCO2
}
CE_er_iw=data.frame(CE_er_iw)#矩阵转化成数据框才能加列名
names(CE_er_iw)=c("Region","Year","Value")#添加列名
CE_er_iw[CE_er_iw<0]=0
write.xlsx(CE_er_iw,file="OUT/CarbonEmission.xlsx",sheetName="CE_er_iw",append=TRUE)


########################## 1.2 工业废物回收的能源（选用E_er_iw）#########################
###########本来就不包含非生物质############
E_er_iw=matrix(0,1770,3)
for (n in 1:30){
  F32=subset(Flows,Region==RG[n,2]&Flow=="F32")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  E_er_iw[(59*n-58):(59*n),1]=RG[n,2]
  E_er_iw[(59*n-58):(59*n),2]=1961:2019
  F32[is.na(F32)]=0
  E_er_iw[(59*n-58):(59*n),3]=F32$Value*1000*0.0156#最终单位GJ.假设和柴木热量一致
}
E_er_iw=data.frame(E_er_iw)#矩阵转化成数据框才能加列名
names(E_er_iw)=c("Region","Year","Value")#添加列名
E_er_iw[E_er_iw<0]=0
write.xlsx(E_er_iw,file="OUT/CarbonEmission.xlsx",sheetName="E_er_iw",append=TRUE)

########################## 1.3 工业废物能源回收避免的碳排放（选用AE_er_iw）#########################
##本来就不包含非生物质
AE_er_iw=matrix(0,1770,3)
for (n in 1:30){
  F32=subset(Flows,Region==RG[n,2]&Flow=="F32")
  ELC_CI_n=subset(ELC_CI,Region==RG[n,2])
  AE_er_iw[(59*n-58):(59*n),1]=RG[n,2]
  AE_er_iw[(59*n-58):(59*n),2]=1961:2019
  AE_er_iw[(59*n-58):(59*n),3]=F32$Value*1000*0.0156*0.25*ELC_CI_n$CI/1000#最终单位为kg CO2
}
AE_er_iw=data.frame(AE_er_iw)#矩阵转化成数据框才能加列名
names(AE_er_iw)=c("Region","Year","Value")#添加列名
AE_er_iw[AE_er_iw<0]=0
write.xlsx(AE_er_iw,file="OUT/CarbonEmission.xlsx",sheetName="AE_er_iw",append=TRUE)#导出数据24

##原本的非能源回收和填埋目前只针对非生物质了，没有碳排放和碳存量，放到“废弃代码备份.R”脚本中了

################################ （七）生产过程能耗与碳排放 ####################################################
################################ 1. 生产过程直接能耗(自下而上：单耗*产量) #########################################
Product_P_E=dcast(Product_P,Region+Year~Product,value.var = "Value")
Product_P_E$Printing=Product_P_E$Newsprint+Product_P_E$`Printing and writing papers`+Product_P_E$`Packaging paper and paperboard`*0.17
Product_P_E=melt(Product_P_E,id=c("Region","Year"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
Product_P_E=Product_P_E[order(Product_P_E$Product), ]
Product_P_E=Product_P_E[order(Product_P_E$Region), ]
Product_P_E$Year=1961:2019
Product_P_E=Product_P_E[order(Product_P_E$Year), ]
Product_P_W=dcast(Product_P_E,Region+Year~Product,value.var = "Value")
write.xlsx(Product_P_W,file="OUT/Product_Data.xlsx",sheetName="Product_P_W",append=TRUE)#构造一个工作簿，放置生产与消费数据


SEC_prd_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Production_SEC")
SEC_prd_W=SEC_prd_W[,1:12]
SEC_prd=melt(SEC_prd_W,id=c("Region","Year"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
SEC_prd=SEC_prd[order(SEC_prd$Region), ]
SEC_prd=SEC_prd[order(SEC_prd$Year), ]
SEC_prd=subset(SEC_prd,Region!="World")
E_prd=SEC_prd[,1:3]
E_prd$Value=Product_P_E$Value*SEC_prd$Value/1000000000#单位为GJ
write.xlsx(E_prd,file="OUT/CarbonEmission.xlsx",sheetName="E_prd",append=TRUE)#E_prd是生产阶段带有产品结构的直接能耗

############################################### 2. 总的直接能耗 ##############################################################
E_prd_W=dcast(E_prd,Region+Year~Product,value.var = "Value")
E_prd_W$Total=E_prd_W$Chemical.wood.pulp+E_prd_W$Household.and.sanitary.papers+E_prd_W$Mechanical.wood.pulp+
  E_prd_W$Newsprint+E_prd_W$Other.paper.and.paperboard+E_prd_W$Packaging.paper.and.paperboard+
  E_prd_W$Printing.and.writing.papers+E_prd_W$Pulp.from.fibres.other.than.wood+E_prd_W$Recycled.pulp+E_prd_W$Printing
E_prd_P=E_prd_W[,-(3:12)]
names(E_prd_P)=c("Region","Year","Value")
write.xlsx(E_prd_P,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_P",append=TRUE)#E_prd_P是生产阶段总的直接能耗

############################################# 3. 估算生物质能源 #####################################################################
##############IEA能源结构计算得到的生物质能源VS通过备料损失率和工业废物能源回收推算得到的生物质能源########
############################################# 3.1 IEA能源结构计算得到的生物质能源(初级能源)：上面分能源类型的初级能耗的生物质相关能耗汇总得到 #####
E_prd_Pri_IEA=read.xlsx(file="INPUT/Energy structure and emissions.xlsx",sheetName="TCV1960")
E_prd_Pri_IEA=subset(E_prd_Pri_IEA,Region!="World")
E_prd_Pri_Bio_IEA=E_prd_Pri_IEA[,1:2]
E_prd_Pri_Bio_IEA$Value=rowSums(E_prd_Pri_IEA[,28:36])
E_prd_Pri_Bio_IEA$Value=E_prd_Pri_Bio_IEA$Value*1000
write.xlsx(E_prd_Pri_Bio_IEA,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri_Bio_IEA",append=TRUE)#E_prd_BioWst是IEA能源结构推算得到的生物质供能（初级）

############################################ 3.2 通过备料损失率和工业废物能源回收推算得到的生物质能源(初级能源) ##################################
F1=subset(Flows,Flow=="F1")
F1=F1[order(F1$Region), ]
F2=subset(Flows,Flow=="F2")
F2=F2[order(F2$Region), ]
F3=subset(Flows,Flow=="F3")                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
F3=F3[order(F3$Region), ]
Wood=F1[,-(2:4)]
Wood=Wood[,-3]
Wood$Value=F1$Value+F2$Value
Otherfibre=F3[,-(2:4)]
Otherfibre=Otherfibre[,-3]
Otherfibre$Value=F3$Value
E_er_iw=read.xlsx(file="OUT/CarbonEmission.xlsx",sheetName="E_er_iw")#导入工厂能源回收量
E_er_iw_EW=subset(E_er_iw,Region!=0)
E_prd_BioWst_estm=E_er_iw_EW[,2:3]
E_prd_BioWst_estm$Value=Wood$Value*0.056*15.6/(1-0.056)+Otherfibre$Value*0.056*14.7/(1-0.056)+as.numeric(E_er_iw_EW$Value)#estm尾缀代表估算值、推算值;0.056是损失率，15.6GJ/t
##计算木材边角料供能
E_prd_BioWst_estm$WoodE=Wood$Value*0.056*15.6/(1-0.056)+Otherfibre$Value*0.056*14.7/(1-0.056)#14.7GJ/Ton来自中国能源统计年鉴各种能源折标煤参考系数，是四种秸秆的均值
E_prd_BioWst_estm$iwer=as.numeric(E_er_iw_EW$Value)
write.xlsx(E_prd_BioWst_estm,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri_BioWst_estm",append=TRUE)#E_prd_BioWst_estm是估算的生物质能源（初级）供应量

########################################### 3.3 两组生物质初级能源消耗量的对比 ############################################
##构造表格E_prd_Bio_Cmpr，输入IEA的生物质供能数据以及估算的生物质供能数据
E_prd_Bio_Cmpr=E_prd_BioWst_estm[,1:2]
E_prd_Bio_Cmpr$IEA=E_prd_Pri_Bio_IEA$Value
E_prd_Bio_Cmpr$Estm=E_prd_BioWst_estm$Value
E_prd_Bio_Cmpr_L=melt(E_prd_Bio_Cmpr,id=c("Year","Region"),variable.name = "Item",value.name = "Value")
write.xlsx(E_prd_Bio_Cmpr_L,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Bio_Cmpr",append=TRUE)
##导入数据
E_prd_Bio_Cmpr=read.xlsx(file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Bio_Cmpr")
##绘制折线图
p=ggplot(subset(E_prd_Bio_Cmpr,Region!="World"), mapping = aes(x = Year, y = Value/1000000, group=Item,colour=Item)) + geom_line(size=0.8,aes())+
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
  labs(text=element_text(family='calibri'),y="Bio-Energy consumption: PJ",x='Year',title='Bio-Energy Consumption Comparison',colour="Bio-Energy consumption")+
  facet_wrap(~Region,ncol=5,scales = "free")+
  scale_x_discrete(breaks=seq(1960, 2020, 20))
p
ggsave(p, file="OUT/Plot/EnergyConsumption/BioEComp.jpeg", width=12, height=8)

############################################### 4. 分能源类型结构的直接能耗 ###################################################
E_prd_Bio=E_prd_Pri_Bio_IEA
E_prd_Bio$Value=E_prd_BioWst_estm$Value*0.85
E_prd_ExBio=E_prd_Bio
E_prd_ExBio$Value=E_prd_P$Value-E_prd_Bio$Value
E_prd_ExBio[E_prd_ExBio<0]
ES_ExBio=read.xlsx("INPUT/Energy structure and emissions.xlsx",sheetName = "TCV1960_%_EB")#ES_ExBio是基于IEA数据计算的、除去生物质能源的直接能耗的能源结构比例；埃及和马来西亚单独找了工业层面的能源消费结构数据补充上的
E_prd_ExBio_ES=E_prd_Bio
E_prd_ExBio_ES[,3:32]=E_prd_ExBio$Value*ES_ExBio[,3:32]
E_prd_ES=E_prd_ExBio_ES
E_prd_ES$Wood=E_prd_BioWst_estm$WoodE*0.85
E_prd_ES$Black.liquid=E_prd_BioWst_estm$iwer*0.85
write.xlsx(E_prd_ES,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_ES",append=TRUE)#E_prd_ES是分能源类型结构的直接能耗

############################################### 5. 分能源类型结构的初级能耗 ######################################
ETF=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "ETF")#ETF是分能源类型的转换效率
ETF=subset(ETF,Region!="World")
E_prd_ES_Pri=E_prd_ES[,3:32]/ETF[,3:32]
E_prd_ES_Pri=cbind(E_prd_ES[,1:2],E_prd_ES_Pri)#加上年份和国家
write.xlsx(E_prd_ES_Pri,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_ES_Pri",append=TRUE)#E_prd_ES_Pri是分能源类型结构的初级能耗


############################################## 6. 总的初级能耗 #################################
E_prd_Pri=E_prd_ES_Pri[,1:2]
E_prd_Pri$Value=rowSums(E_prd_ES_Pri[,3:32])
write.xlsx(E_prd_Pri,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri",append=TRUE)#E_prd_Pri是IEA能源结构推算得到的总初级能耗

#################################### 7. 构造一个表格，将生产过程的总初级能耗按照生物质和非生物质分开#######################
##该表格在CarbonEmission表格中构建一个sheet，名字为E_prd_BioOrNot_AllEstm
E_prd_BioOrNot_AllEstm=E_prd_BioWst_estm[,1:2]
E_prd_BioOrNot_AllEstm$E_prd_Bio=E_prd_BioWst_estm$Value
E_prd_BioOrNot_AllEstm$E_prd_ExBio=E_prd_Pri$Value-E_prd_BioOrNot_AllEstm$E_prd_Bio
write.xlsx(E_prd_BioOrNot_AllEstm,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_BioOrNot_AllEstm_Pri",append=TRUE)

############################# 8. 分配到产品上，并且区分是否生物质能 ##############################
############################# 8.1 计算每种产品能源消耗结构 #####################################
##想把初级能耗和后续碳排分配到每种产品，则应按照每种产品或工艺能源消耗的结构去分配，因此首先应计算该比例
E_prd=read.xlsx(file="OUT/CarbonEmission.xlsx",sheetName="E_prd")
E_prd=E_prd[,-1]#去掉没用的序号
E_prd_W=dcast(E_prd,Region+Year~Product,value.var = "Value")
E_prd_W$SUM=rowSums(E_prd_W[,3:12])
E_prd_PS=cbind(E_prd_W[,1:2],E_prd_W[,3:12]/E_prd_W$SUM)##E_prd_PS为能源消耗按照产品种类分配的比例
write.xlsx(E_prd_PS,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_PS%",append=TRUE)#导出该比例数据

############################# 8.2 分产品结构和分是否生物质能的初级能耗 ######################################
##先按照产品能耗，计算分产品结构的初级能耗
E_prd_PS=read.xlsx(file="OUT/CarbonEmission.xlsx",sheetName="E_prd_PS%")#读入该比例数据
E_prd_PS=E_prd_PS[,-1]#去掉序号
E_prd_Pri_PS=E_prd_Pri$Value*E_prd_PS[,3:12]
E_prd_Pri_PS=cbind(E_prd_Pri[,1:2],E_prd_Pri_PS)
write.xlsx(E_prd_Pri_PS,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri_PS",append=TRUE)#E_prd_Pri_PS是分产品种类结构的初级能耗
##在CarbonEmission工作簿中构建名为E_prd_Pri_PS_BioOrNot的sheet，首先从E_prd_BioOrNot_AllEstm中将E_prd_Bio和E_prd_ExBio两列数据粘贴到E_prd_Pri_PS_BioOrNot的M和N列中
##计算出除了CWP、NWP和PR的其他过程初级能耗比例，备用
E_prd_Pri_PS_pct=E_prd_Pri_PS
E_prd_Pri_PS_pct[,3:12]=E_prd_Pri_PS[,3:12]/
  (rowSums(E_prd_Pri_PS[,3:12])-E_prd_Pri_PS$Chemical.wood.pulp-E_prd_Pri_PS$Printing-
     E_prd_Pri_PS$Pulp.from.fibres.other.than.wood)
E_prd_Pri_PS_pct$Chemical.wood.pulp=0
E_prd_Pri_PS_pct$Printing=0
E_prd_Pri_PS_pct$Pulp.from.fibres.other.than.wood=0
E_prd_Pri_PS_pct[is.na(E_prd_Pri_PS_pct)]=0
##首先确定CWP_Bio和NWP_Bio,然后将生物质能分配给其他过程，主要基于以下两条原则：
##①当生物质供能小于化学制浆能耗的85%时，则将生物质能耗按照两种化学制浆能耗比例分配给两种方法
##②当生物质供能大于化学制浆能耗的85%时，则最多给配给两种化学制浆85%的生物质能源，其余平均分给其他过程
E_prd_Pri_PS_BioOrNot=E_prd_Pri_PS_pct[,1:2]
E_prd_BioOrNot_AllEstm$E_prd_Bio[is.na(E_prd_BioOrNot_AllEstm$E_prd_Bio)]=0
E_prd_Pri_PS[is.na(E_prd_Pri_PS)]=0
for (n in 1:1770){
  if(E_prd_BioOrNot_AllEstm$E_prd_Bio[n]-(E_prd_Pri_PS$Chemical.wood.pulp[n]+E_prd_Pri_PS$Pulp.from.fibres.other.than.wood[n])*0.85<0){
    E_prd_Pri_PS_BioOrNot$CWP_Bio[n]=E_prd_BioOrNot_AllEstm$E_prd_Bio[n]*E_prd_Pri_PS$Chemical.wood.pulp[n]/
      (E_prd_Pri_PS$Chemical.wood.pulp[n]+E_prd_Pri_PS$Pulp.from.fibres.other.than.wood[n])
    E_prd_Pri_PS_BioOrNot$NWP_Bio[n]=E_prd_BioOrNot_AllEstm$E_prd_Bio[n]*E_prd_Pri_PS$Pulp.from.fibres.other.than.wood[n]/
      (E_prd_Pri_PS$Chemical.wood.pulp[n]+E_prd_Pri_PS$Pulp.from.fibres.other.than.wood[n])
  }else{
    E_prd_Pri_PS_BioOrNot$CWP_Bio[n]=E_prd_Pri_PS$Chemical.wood.pulp[n]*0.85
    E_prd_Pri_PS_BioOrNot$NWP_Bio[n]=E_prd_Pri_PS$Pulp.from.fibres.other.than.wood[n]*0.85
  }
}
E_prd_Pri_PS_BioOrNot$HS_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Household.and.sanitary.papers
E_prd_Pri_PS_BioOrNot$MP_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Mechanical.wood.pulp
E_prd_Pri_PS_BioOrNot$NP_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Newsprint
E_prd_Pri_PS_BioOrNot$OP_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Other.paper.and.paperboard
E_prd_Pri_PS_BioOrNot$PP_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Packaging.paper.and.paperboard
E_prd_Pri_PS_BioOrNot$PR_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Printing
E_prd_Pri_PS_BioOrNot$PW_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Printing.and.writing.papers
E_prd_Pri_PS_BioOrNot$RP_Bio=(E_prd_BioOrNot_AllEstm$E_prd_Bio-E_prd_Pri_PS_BioOrNot$CWP_Bio-E_prd_Pri_PS_BioOrNot$NWP_Bio)*E_prd_Pri_PS_pct$Recycled.pulp
##分配非生物质能
E_prd_Pri_PS_BioOrNot$CWP_ExBio=E_prd_Pri_PS$Chemical.wood.pulp-E_prd_Pri_PS_BioOrNot$CWP_Bio
E_prd_Pri_PS_BioOrNot$NWP_ExBio=E_prd_Pri_PS$Pulp.from.fibres.other.than.wood-E_prd_Pri_PS_BioOrNot$NWP_Bio
E_prd_Pri_PS_BioOrNot$HS_ExBio=E_prd_Pri_PS$Household.and.sanitary.papers-E_prd_Pri_PS_BioOrNot$HS_Bio
E_prd_Pri_PS_BioOrNot$MP_ExBio=E_prd_Pri_PS$Mechanical.wood.pulp-E_prd_Pri_PS_BioOrNot$MP_Bio
E_prd_Pri_PS_BioOrNot$NP_ExBio=E_prd_Pri_PS$Newsprint-E_prd_Pri_PS_BioOrNot$NP_Bio
E_prd_Pri_PS_BioOrNot$OP_ExBio=E_prd_Pri_PS$Other.paper.and.paperboard-E_prd_Pri_PS_BioOrNot$OP_Bio
E_prd_Pri_PS_BioOrNot$PP_ExBio=E_prd_Pri_PS$Packaging.paper.and.paperboard-E_prd_Pri_PS_BioOrNot$PP_Bio
E_prd_Pri_PS_BioOrNot$PR_ExBio=E_prd_Pri_PS$Printing-E_prd_Pri_PS_BioOrNot$PR_Bio
E_prd_Pri_PS_BioOrNot$PW_ExBio=E_prd_Pri_PS$Printing.and.writing.papers-E_prd_Pri_PS_BioOrNot$PW_Bio
E_prd_Pri_PS_BioOrNot$RP_ExBio=E_prd_Pri_PS$Recycled.pulp-E_prd_Pri_PS_BioOrNot$RP_Bio
write.xlsx(E_prd_Pri_PS_BioOrNot,file="OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri_PS_BioOrNot",append=TRUE)
############################# 9. 生产过程——碳排(自下而上) ####################################################
##导入各能源二氧化碳含量
CC=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "CO2_Con")
CC=subset(CC,Region!="World")
##计算生产过程能源相关碳排放
CE_prd=E_prd_ES_Pri[,3:32]*CC[,3:32]/1000#单位是kg CO2
CE_prd=cbind(E_prd_ES_Pri[,1:2],CE_prd)
names(CE_prd)=c(colnames(E_prd_ES_Pri))##这一句不加也可以，原本列名就是正确的
CE_prd$SUM=rowSums(CE_prd[,3:32])#各列求和
CE_prd$Bio=rowSums(CE_prd[,28:29])#生物质能源碳排放（包括工业废物、生活废物、生物质燃料）
CE_prd$ExBio=CE_prd$SUM-CE_prd$Bio#化石能源碳排放
write.xlsx(CE_prd,file="OUT/CarbonEmission.xlsx",sheetName="CE_prd",append=TRUE)#分能源结构的碳排放（基于初级能耗）

##以下将总量、生物质和非生物质碳排分别存入表格备用
CE_prd_SUM=CE_prd[,1:2]
CE_prd_SUM$Value=CE_prd$SUM
write.xlsx(CE_prd_SUM,file="OUT/CarbonEmission.xlsx",sheetName="CE_prd_SUM",append=TRUE)#排放总值

CE_prd_Bio=CE_prd[,1:2]
CE_prd_Bio$Value=CE_prd$Bio
write.xlsx(CE_prd_Bio,file="OUT/CarbonEmission.xlsx",sheetName="CE_prd_Bio",append=TRUE)#生物质能源排放

CE_prd_ExBio=CE_prd[,1:2]
CE_prd_ExBio$Value=CE_prd$ExBio
write.xlsx(CE_prd_ExBio,file="OUT/CarbonEmission.xlsx",sheetName="CE_prd_ExBio",append=TRUE)#非生物质能源排放

##【说明】：
##假设生物质能源回收只供应给化学制浆和非木材化学制浆，其他制浆造纸过程能源都来自于化石能源
##于是，在CarbonEmission表格中构建一张sheet，名为E_prd_Pri_PS_BioOrNot，既区分产品又区分能源结构
##构建另外一个名为CE_prd_PS_BioOrNot的sheet，存放既区分产品又区分能源结构的生产过程能源消费碳排放数据，具体处理方法见excel中的公式##
E_prd_Pri_PS_BioOrNot=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = "E_prd_Pri_PS_BioOrNot")
E_prd_Pri_PS_BioOrNot=E_prd_Pri_PS_BioOrNot[,-1]
CE_prd_PS_BioOrNot=E_prd_Pri_PS_BioOrNot
CE_prd_Bio=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = "CE_prd_Bio")
CE_prd_Bio=CE_prd_Bio[,-1]
CE_prd_ExBio=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = "CE_prd_ExBio")
CE_prd_ExBio=CE_prd_ExBio[,-1]
CE_prd_PS_BioOrNot[,3:12]=CE_prd_Bio$Value*E_prd_Pri_PS_BioOrNot[,3:12]/rowSums(E_prd_Pri_PS_BioOrNot[,3:12])
CE_prd_PS_BioOrNot[,13:22]=CE_prd_ExBio$Value*E_prd_Pri_PS_BioOrNot[,13:22]/rowSums(E_prd_Pri_PS_BioOrNot[,13:22])
CE_prd_PS_BioOrNot[,23:32]=CE_prd_PS_BioOrNot[,3:12]+CE_prd_PS_BioOrNot[,13:22]
CE_prd_PS_BioOrNot[is.na(CE_prd_PS_BioOrNot)]=0
write.xlsx(CE_prd_PS_BioOrNot,file="OUT/CarbonEmission.xlsx",sheetName="CE_prd_PS_BioOrNot",append=TRUE)#非生物质能源排放
##备注：该表格第23:32列，列名有点问题，应该去掉后面的下划线和Bio。



############################## （八）工厂废物处理2 #######################
############################## 0. 比较特殊：边角料能源回收 ############
############################## 0.1 边角料产生的碳排 ############
E_prd_Pri_BioWst_estm=read.xlsx("OUT/CarbonEmission.xlsx",sheetName="E_prd_Pri_BioWst_estm")
CE_WoodE=E_prd_Pri_BioWst_estm
CE_WoodE=CE_WoodE[,-1]
CE_WoodE=CE_WoodE[,-(3:6)]
CE_prd_Bio=read.xlsx("OUT/CarbonEmission.xlsx",sheetName="CE_prd_Bio")
CE_er_iw=CE_prd$Black.liquid
CE_WoodE$Value=CE_prd$Wood
write.xlsx(CE_WoodE,file="OUT/CarbonEmission.xlsx",sheetName="CE_WoodE",append=TRUE)

############################## 0.2 边角料回收的能源#########################
E_WoodE=CE_WoodE
E_WoodE=E_WoodE[,-3]
E_WoodE$Value=CE_WoodE$Value*(12*0.0147)/(0.45*44)##边角料低位热值为0.0147
write.xlsx(E_WoodE,file="OUT/CarbonEmission.xlsx",sheetName="E_WoodE",append=TRUE)
##已经用于本系统了，所以该能源避免的碳排放不再计入最终结果

############################## 0.3 边角料能源回收避免的碳排放#########################
###本来就不包含非生物质
ELC_CI=ELC_CI[order(ELC_CI$Region), ]
ELC_CI=subset(ELC_CI,Region!="World")
AE_WoodE=E_WoodE
AE_WoodE=AE_WoodE[,-3]
AE_WoodE$Value=E_WoodE$Value*0.25*ELC_CI$CI/1000##0.25是转化率
write.xlsx(AE_WoodE,file="OUT/CarbonEmission.xlsx",sheetName="AE_WoodE",append=TRUE)#导出数据24
##已经用于本系统了，所以该能源避免的碳排放不再计入最终结果

############################## 三、绘图前数据准备 ###################################################
##############################（一）分产品的阶段，长数据转宽数据，求和 #####################
##reshape2中的dcast函数可以完成数据长转宽的需求
TN1=c( "CE_chem","CE_wttrt","CS_ps_Bio","E_chem")
for (i in 1:4){
  data0=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = TN1[i])
  data0$Value=as.numeric(data0$Value)
  data_P=dcast(data0,Region+Year~Product,value.var = "Value")
  NC=ncol(data_P)
  data_PA=data_P[,3:NC]
  data_T=data_P[,1:2]
  data_T$Value=apply(data_PA,1,sum)
  write.xlsx(data_T,file="OUT/CarbonEmission.xlsx",sheetName=paste(TN1[i],"_P",sep = "",collapse = NULL),append=TRUE)
}

############################## （二）生命周期能耗表格构造 ##########################################
#在构造以下表格之前，先确保各表格地区和年份排序一致
E_prd_Pri=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = "E_prd_Pri")
E_prd_Pri=E_prd_Pri[,-1]
TN2=c("E_chem_P","E_he","E_nwh", "E_rpc","E_wh","E_prd_Pri")
Energy=matrix(0,1770,6)
Energy=data.frame(Energy)
Energy[,1:2]=E_prd_Pri[,1:2]
    for (i in 1:6){
      Energy_i=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = TN2[i])
      Energy_i=subset(Energy_i,Region!="World")
      Energy_i=Energy_i[order(Energy_i$Region), ]
      Energy[,i+2]=as.numeric(Energy_i$Value)
    }
names(Energy)=c("Region","Year","E_chem_P","E_he","E_nwh", "E_rpc","E_wh","E_prd_Pri")
write.xlsx(Energy,file="OUT/EnergySheet.xlsx",sheetName="EnergySheet")

############################## （三）碳排放数据整理 ###################################################
############################## 1. 生命周期碳排放和碳汇构造表格 ###################################################
TN2=c("CS_pw","CS_nw","CS_ps_Bio_SUM", "CS_lf_mw_Bio","CS_ner_mw_Bio",
      "AE_er_mw","AE_lf_mw_Bio","AE_er_iw","AE_WoodE",
      "CE_df","CE_wh","CE_nwh","CE_rpc","CE_chem_P","CE_wttrt_Pulp","CE_wttrt_Paper","CE_wttrt_ToNature_Pulp","CE_wttrt_ToNature_Paper",
      "CE_prd_ExBio","CE_prd_Bio",
      "CE_inc_Bio","CE_er_mw_Bio", "CO2_lf_mw_Bio","CE_lf_mw_Bio","CECP_lf_mw_Bio","CESC_lf_mw_Bio",
      "CE_er_iw","CE_WoodE")
CE_prd=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = "CE_prd")
Carbon=matrix(0,1770,25)
Carbon=data.frame(Carbon)
Carbon[,1:2]=CE_prd[,2:3]

for (i in 1:9){
  Carbon_i=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = TN2[i])
  Carbon_i=Carbon_i[order(Carbon_i$Region), ]
  Carbon_i=subset(Carbon_i,Region!="World"&Region!=0)
  Carbon[,i+2]=-as.numeric(Carbon_i$Value)
}

for (i in 10:28){
  Carbon_i=read.xlsx("OUT/CarbonEmission.xlsx",sheetName = TN2[i])
  Carbon_i=Carbon_i[order(Carbon_i$Region), ]
  Carbon_i=subset(Carbon_i,Region!="World"&Region!=0)
  Carbon[,i+2]=as.numeric(Carbon_i$Value)
}
names(Carbon)=c("Year","Region","CS_pw","CS_nw","CS_ps_Bio_SUM", "CS_lf_mw_Bio","CS_ner_mw_Bio",
                "AE_er_mw","AE_lf_mw_Bio","AE_er_iw","AE_WoodE",
                "CE_df","CE_wh","CE_nwh","CE_rpc","CE_chem_P","CE_wttrt_Pulp","CE_wttrt_Paper","CE_wttrt_ToNature_Pulp","CE_wttrt_ToNature_Paper",
                "CE_prd_ExBio","CE_prd_Bio",
                "CE_inc_Bio","CE_er_mw_Bio", "CO2_lf_mw_Bio","CE_lf_mw_Bio","CECP_lf_mw_Bio","CESC_lf_mw_Bio",
                "CE_er_iw","CE_WoodE")
Carbon[is.na(Carbon)]=0##将俄罗斯1992年前的缺失值改为0
#write.xlsx(Carbon,file="OUT/CarbonSheet.xlsx",sheetName="CarbonSheet")
write.xlsx(Carbon,file="OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName="CarbonSheet")

############################## 2. 重分类、绘图前准备等处理 ###################################################
############################## 2.1 所有来源的碳排放都计入 ###################################################
############################## 2.1.1 重分类 ###################################################################
##根据分析需求，对碳汇和碳源重分类
CarbonSheet=read.xlsx(file="OUT/CarbonSheet.xlsx",sheetName="CarbonSheet")
CarbonSheet=CarbonSheet[,-1]##把第一列没用的序号去掉
Carbon_ReCls= CarbonSheet[,1:2]##定义前两列地区和年份
##定义其后每一列数据
Carbon_ReCls$CS_ps=CarbonSheet$CS_ps_Bio_SUM
Carbon_ReCls$CS_lf=CarbonSheet$CS_lf_mw_Bio
Carbon_ReCls$CS_ner=CarbonSheet$CS_ner_mw_Bio-CarbonSheet$CE_wttrt_ToNature_Pulp-CarbonSheet$CE_wttrt_ToNature_Paper
Carbon_ReCls$AE_er=CarbonSheet$AE_er_mw
Carbon_ReCls$AE_lf=CarbonSheet$AE_lf_mw_Bio
Carbon_ReCls$CE_df=CarbonSheet$CE_df
Carbon_ReCls$CE_he=CarbonSheet$CE_wh+CarbonSheet$CE_nwh+CarbonSheet$CE_rpc
Carbon_ReCls$CE_chem=CarbonSheet$CE_chem_P
Carbon_ReCls$CE_wttrt=CarbonSheet$CE_wttrt_Pulp+CarbonSheet$CE_wttrt_Paper
Carbon_ReCls$CE_prd_EXBio=CarbonSheet$CE_prd_ExBio
Carbon_ReCls$CE_inc=CarbonSheet$CE_inc_Bio
Carbon_ReCls$CE_er=CarbonSheet$CE_er_mw_Bio+CarbonSheet$CE_er_iw+CarbonSheet$CE_WoodE
Carbon_ReCls$CE_lf=CarbonSheet$CE_lf_mw_Bio
Carbon_ReCls$CECP_lf=CarbonSheet$CECP_lf_mw_Bio
Carbon_ReCls$CESC_lf=CarbonSheet$CO2_lf_mw_Bio+CarbonSheet$CESC_lf_mw_Bio
##增加三列，分别是碳汇、碳排和净碳排
Carbon_ReCls$CS=rowSums(Carbon_ReCls[,3:7])/10^9##CE表示Carbon Sink or Carbon Stock
Carbon_ReCls$CE=rowSums(Carbon_ReCls[,8:17])/10^9##CE表示是Carbon Emission
Carbon_ReCls$NCE=Carbon_ReCls$CE+Carbon_ReCls$CS##NCE表示Net Carbon Emission
Carbon_ReCls[is.na(Carbon_ReCls)]=0
Carbon_ReCls_YearSum=matrix(0,30,ncol(Carbon_ReCls))
for (n in 1:30){
  Carbon_ReCls_YearSum[n,1]="1961_2019"
  Carbon_ReCls_YearSum[n,2]=RG[n,2]
  Carbon_ReCls_YearSum[n,3:ncol(Carbon_ReCls)]=as.numeric(colSums(Carbon_ReCls[(59*n-58):(59*n),3:ncol(Carbon_ReCls)]))
}
Carbon_ReCls_YearSum=data.frame(Carbon_ReCls_YearSum)##转换成数据框
names(Carbon_ReCls_YearSum)=colnames(Carbon_ReCls)
Carbon_ReCls=rbind(Carbon_ReCls,Carbon_ReCls_YearSum)##将累积数据添加进去
Carbon_ReCls[,3:ncol(Carbon_ReCls)]=as.numeric(unlist(Carbon_ReCls[,3:ncol(Carbon_ReCls)]))
##导出数据
write.xlsx(Carbon_ReCls,file="OUT/CarbonSheet.xlsx",sheetName="Carbon_ReCls",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】


######################################## 2.1.2 取出4个时间节点的净碳排放数据 ############################################################
Carbon_ReCls=read.xlsx(file="OUT/CarbonSheet.xlsx",sheetName="Carbon_ReCls")##因为表格有手动处理，因此重新导入
Carbon_ReCls=Carbon_ReCls[,-1]##去掉无用序号列
LcNCE_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcNCE_4Years=LcNCE_4Years[,-(3:19)]
LcNCE_4Years=dcast(LcNCE_4Years,Region~Year,value.var = "NCE")
write.xlsx(LcNCE_4Years,file="OUT/CarbonSheet.xlsx",sheetName="LcNCE_4Years",append = TRUE)##导出选定4年的净碳排数据

######################################## 2.1.3 计算碳存储量各部分占比 ###########################################################
LcCS_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcCS_4Years=LcCS_4Years[,-(8:17)]
LcCS_4Years=LcCS_4Years[,-(9:10)]
LcCS_4Years[is.na(LcCS_4Years)]=0
LcCS_4Years_Strc=LcCS_4Years[,-8]##LcCS_4Years_Strc表示 structure of lifecycle carbon sinks of 4 years
LcCS_4Years_Strc[,3:7]=as.numeric(unlist(LcCS_4Years[,3:7]))/(as.numeric(unlist(LcCS_4Years[,8]))*10^9)
LcCS_4Years_Strc$CS=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")$CS##加一列：负碳排放的总量
LcCS_4Years_Strc[is.na(LcCS_4Years_Strc)]=0
write.xlsx(LcCS_4Years_Strc,file="OUT/CarbonSheet.xlsx",sheetName="LcCS_4Years_Strc",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】


######################################## 2.1.4 计算碳排放量各部分占比 ###########################################################
LcCE_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcCE_4Years=LcCE_4Years[,-(20)]
LcCE_4Years=LcCE_4Years[,-(18)]
LcCE_4Years=LcCE_4Years[,-(3:7)]
LcCE_4Years[is.na(LcCE_4Years)]=0
LcCE_4Years_Strc=LcCE_4Years[,-13]##LcCE_4Years_Strc表示 structure of lifecycle carbon emissions of 4 years
LcCE_4Years_Strc[,3:12]=as.numeric(unlist(LcCE_4Years[,3:12]))/(as.numeric(unlist(LcCE_4Years[,13]))*10^9)
LcCE_4Years_Strc$CE=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")$CE##加一列：碳排放的总量
LcCE_4Years_Strc[is.na(LcCE_4Years_Strc)]=0
write.xlsx(LcCE_4Years_Strc,file="OUT/CarbonSheet.xlsx",sheetName="LcCE_4Years_Strc",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】


############################## 2.2 生物质来源的碳排放视为碳中性 ###################################################
##建立CarbonSheet工作簿的副本，改名为CarbonSheet_BioCarbonNeutral
##删除重分类之后的四个sheet
############################## 2.2.1 重分类 ###################################################################
##根据分析需求，对碳汇和碳源重分类
CarbonSheet=read.xlsx(file="OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName="CarbonSheet")
CarbonSheet=CarbonSheet[,-1]##把第一列没用的序号去掉
Carbon_ReCls= CarbonSheet[,1:2]##定义前两列地区和年份
##定义其后每一列数据
Carbon_ReCls$CS_ps=CarbonSheet$CS_ps_Bio_SUM
Carbon_ReCls$CS_lf=CarbonSheet$CS_lf_mw_Bio
Carbon_ReCls$CS_ner=CarbonSheet$CS_ner_mw_Bio-CarbonSheet$CE_wttrt_ToNature_Pulp-CarbonSheet$CE_wttrt_ToNature_Paper
Carbon_ReCls$AE_er=CarbonSheet$AE_er_mw
Carbon_ReCls$AE_lf=CarbonSheet$AE_lf_mw_Bio
Carbon_ReCls$CE_df=CarbonSheet$CE_df
Carbon_ReCls$CE_he=CarbonSheet$CE_wh+CarbonSheet$CE_nwh+CarbonSheet$CE_rpc
Carbon_ReCls$CE_chem=CarbonSheet$CE_chem_P
Carbon_ReCls$CE_wttrt=CarbonSheet$CE_wttrt_Pulp+CarbonSheet$CE_wttrt_Paper
Carbon_ReCls$CE_prd_EXBio=CarbonSheet$CE_prd_ExBio
Carbon_ReCls$CE_inc=CarbonSheet$CE_inc_Bio
Carbon_ReCls$CE_er=CarbonSheet$CE_er_mw_Bio+CarbonSheet$CE_er_iw+CarbonSheet$CE_WoodE
Carbon_ReCls$CE_lf=CarbonSheet$CE_lf_mw_Bio
Carbon_ReCls$CECP_lf=CarbonSheet$CECP_lf_mw_Bio
Carbon_ReCls$CESC_lf=CarbonSheet$CO2_lf_mw_Bio+CarbonSheet$CESC_lf_mw_Bio
##增加三列，分别是碳汇、碳排和净碳排
Carbon_ReCls$CS=rowSums(Carbon_ReCls[,3:7])/10^9##CE表示Carbon Sink or Carbon Stock
Carbon_ReCls$CE=(Carbon_ReCls$CE_df+Carbon_ReCls$CE_he+Carbon_ReCls$CE_chem+Carbon_ReCls$CE_wttrt+
  Carbon_ReCls$CE_prd_EXBio+Carbon_ReCls$CE_lf)/10^9##CE表示是Carbon Emission
Carbon_ReCls$NCE=Carbon_ReCls$CE+Carbon_ReCls$CS##NCE表示Net Carbon Emission
Carbon_ReCls[is.na(Carbon_ReCls)]=0##缺失值替换为0

Carbon_ReCls_YearSum=matrix(0,30,ncol(Carbon_ReCls))
for (n in 1:30){
  Carbon_ReCls_YearSum[n,1]="1961_2019"
  Carbon_ReCls_YearSum[n,2]=RG[n,2]
  Carbon_ReCls_YearSum[n,3:ncol(Carbon_ReCls)]=as.numeric(unlist(colSums(Carbon_ReCls[(59*n-58):(59*n),3:ncol(Carbon_ReCls)])))
}
Carbon_ReCls_YearSum=data.frame(Carbon_ReCls_YearSum)##转换成数据框
names(Carbon_ReCls_YearSum)=colnames(Carbon_ReCls)
Carbon_ReCls=rbind(Carbon_ReCls,Carbon_ReCls_YearSum)##将累积数据添加进去
Carbon_ReCls[,3:ncol(Carbon_ReCls)]=as.numeric(unlist(Carbon_ReCls[,3:ncol(Carbon_ReCls)]))
##导出数据
write.xlsx(Carbon_ReCls,file="OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName="Carbon_ReCls",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】



######################################## 2.2.2 取出4个时间节点的净碳排放数据 ############################################################
Carbon_ReCls=read.xlsx(file="OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName="Carbon_ReCls")##因为表格有手动处理，因此重新导入
Carbon_ReCls=Carbon_ReCls[,-1]##去掉无用序号列
LcNCE_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcNCE_4Years=LcNCE_4Years[,-(3:19)]
LcNCE_4Years=dcast(LcNCE_4Years,Region~Year,value.var = "NCE")
write.xlsx(LcNCE_4Years,file="OUT/CarbonSheet_BioCarbonNeutral.xlsx",sheetName="LcNCE_4Years",append = TRUE)##导出选定4年的净碳排数据

######################################## 2.2.3 计算碳存储量各部分占比 ###########################################################
LcCS_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcCS_4Years=LcCS_4Years[,-(8:17)]
LcCS_4Years=LcCS_4Years[,-(9:10)]
LcCS_4Years[is.na(LcCS_4Years)]=0
LcCS_4Years_Strc=LcCS_4Years[,-8]##LcCS_4Years_Strc表示 structure of lifecycle carbon sinks of 4 years
LcCS_4Years_Strc[,3:7]=as.numeric(unlist(LcCS_4Years[,3:7]))/(as.numeric(unlist(LcCS_4Years[,8]))*10^9)
LcCS_4Years_Strc$CS=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")$CS##加一列：负碳排放的总量
LcCS_4Years_Strc[is.na(LcCS_4Years_Strc)]=0
write.xlsx(LcCS_4Years_Strc,file="OUT/CarbonSheet_BioCarbonNeutral1201.xlsx",sheetName="LcCS_4Years_Strc",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】


######################################## 2.2.4 计算碳存排放各部分占比 ###########################################################
LcCE_4Years=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")
LcCE_4Years=LcCE_4Years[,-(20)]
LcCE_4Years=LcCE_4Years[,-(18)]
LcCE_4Years=LcCE_4Years[,-(3:7)]
LcCE_4Years=LcCE_4Years[,-(11:12)]##去掉碳中性的项目
LcCE_4Years=LcCE_4Years[,-(8:9)]##去掉碳中性的项目
LcCE_4Years[is.na(LcCE_4Years)]=0
LcCE_4Years_Strc=LcCE_4Years[,-13]##LcCE_4Years_Strc表示 structure of lifecycle carbon emissions of 4 years
LcCE_4Years_Strc=LcCE_4Years_Strc[,-(11:12)]##去掉碳中性的项目
LcCE_4Years_Strc=LcCE_4Years_Strc[,-(8:9)]##去掉碳中性的项目
LcCE_4Years_Strc[,3:8]=as.numeric(unlist(LcCE_4Years[,3:8]))/(as.numeric(unlist(LcCE_4Years[,9]))*10^9)
LcCE_4Years_Strc$CE=subset(Carbon_ReCls,Year=="1961"|Year=="1980"|Year=="2000"|Year=="2019"|Year=="1961_2019")$CE##加一列：碳排放的总量
LcCE_4Years_Strc[is.na(LcCE_4Years_Strc)]=0
write.xlsx(LcCE_4Years_Strc,file="OUT/CarbonSheet_BioCarbonNeutral1201.xlsx",sheetName="LcCE_4Years_Strc",append = TRUE)##导出数据
##检查一下俄罗斯数据，如果1961和1980两年有缺失值替换成0【理论上经过上述处理应该没有了】



################################## （四）全生命周期的全球多年累积物质流计算 #############
################################## 1. 全部物质 ##########################################
##################################(1) 年份累积#####################################333
Flows=read.xlsx(file="INPUT/Global Paper_Info.xlsx",sheetName="Flows")
write.xlsx(Flows,file="OUT/Flows.xlsx",sheetName="Flows",append = TRUE)
Flows=read.xlsx(file="OUT/Flows.xlsx",sheetName="Flows")
Flows_L=melt(Flows,id=c("Region","Flow","Input","Output"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
Flows_simple=Flows_L[,-(3:4)]
Flows_WProcess=dcast(Flows_simple,Region+Year~Flow,value.var = "Value")
Flows_WProcess$Year=1961:2019
Flows_WProcess[is.na(Flows_WProcess)]=0#缺失值假设为零
Flows_WProcess[sapply(Flows_WProcess,is.infinite)]=0#错误值假设为零

Flows_YearSum=matrix(0,30,46)
for (n in 1:30){
  Flows_YearSum[n,1]=RG[n,2]
  Flows_YearSum[n,2:46]=as.numeric(colSums(Flows_WProcess[(59*n-58):(59*n),3:47]))
}
Flows_YearSum=data.frame(Flows_YearSum)
names(Flows_YearSum)=c(colnames(Flows_WProcess))[-2]
write.xlsx(Flows_YearSum,file="OUT/Flows.xlsx",sheetName="YearSum",append = TRUE)


#################################（2）全球加总###########################
Year=c(1961:2019)
Year=data.frame(Year)
Flows_Global=matrix(0,59,46)
Flows_WProcess=Flows_WProcess[order(Flows_WProcess$Year),]
for (n in 1:59){
  Flows_Global[n,1]=Year[n,1]
  Flows_Global[n,2:46]=as.numeric(colSums(Flows_WProcess[(30*n-29):(30*n),3:47]))
}
Flows_Global=data.frame(Flows_Global)
colnames(Flows_Global)=colnames(Flows_WProcess)[-1]
write.xlsx(Flows_Global,file="OUT/Flows.xlsx",sheetName="Global",append = TRUE)


################################## 2. 部分流量转化为纯生物质 #########################################
##做下面之前，先把Flows_Bio表格里俄罗斯1992年之前数据改为0
Flows_Bio=read.xlsx(file="OUT/Flows.xlsx",sheetName="Flows_Bio")
Flows_Bio=Flows_Bio[,-1]
Flows_Bio=melt(Flows_Bio,id=c("Region","Flow","Input","Output"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
Flows_Bio_simple=Flows_Bio[,-(3:4)]
Flows_Bio_WProcess=dcast(Flows_Bio_simple,Region+Year~Flow,value.var = "Value")
Flows_Bio_WProcess$Year=1961:2019
Flows_Bio_WProcess[is.na(Flows_Bio_WProcess)]=0#缺失值假设为零
Flows_Bio_WProcess[sapply(Flows_Bio_WProcess,is.infinite)]=0#错误值假设为零

Flows_Bio_YearSum=matrix(0,30,19)
for (n in 1:30){
  Flows_Bio_YearSum[n,1]=RG[n,2]
  Flows_Bio_YearSum[n,2:19]=as.numeric(colSums(Flows_Bio_WProcess[(59*n-58):(59*n),3:20]))
}
Flows_Bio_YearSum=data.frame(Flows_Bio_YearSum)
names(Flows_Bio_YearSum)=c(colnames(Flows_Bio_WProcess))[-2]
write.xlsx(Flows_Bio_YearSum,file="OUT/Flows.xlsx",sheetName="FlowsBio_YearSum",append = TRUE)
##注释：上面设置的转换为数值型没有起作用，需要在表格中手动修改成数值型

