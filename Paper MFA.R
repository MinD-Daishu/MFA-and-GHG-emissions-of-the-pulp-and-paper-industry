#Paper MFA
install.packages("xlsx")
install.packages("data.table") 

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

RG=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Regions")
FlowName=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "FlowName")
for (j in 1:30){
  Filename=paste("INPUT/PaperMFA_",RG[j,2],".xlsx",sep = "",collapse = NULL)
  print(Filename)
}

########################################### 1. 首先，计算废水处理过程的碳流 ##############################
##废水处理过程碳排放
Data_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "FAOSTAT")
Data=melt(Data_W,id=c("Region","Product","Item"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
Product_P=subset(Data,(Product=="Mechanical wood pulp"|Product=="Chemical wood pulp"|
                         Product=="Pulp from fibres other than wood"|Product=="Recycled pulp"|
                         Product=="Newsprint"|Product=="Printing and writing papers"|
                         Product=="Household and sanitary papers"|Product=="Packaging paper and paperboard"|
                         Product=="Other paper and paperboard")&Item=="Production"&Region!="World")
##以上数据，Recycled pulp数据待定，本脚本后续会更正
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
CE_wttrt=COD[,1:3]##排放到空气的碳
CE_wttrt$Value=Product_P$Value*COD$Value*Removal$Value*MCF$Value*0.25*25#乘以25转化为CO2当量，单位是kg

###排向自然水体的碳
CE_wttrt_ToNature=COD[,1:3]
CE_wttrt_ToNature$Value=((Product_P$Value*COD$Value-19.385)*(1/1.8646)-CE_wttrt$Value*12/(25*16))*44/12#单位是kg CO2

###制浆排放到空气
CE_wttrt_Pulp=subset(CE_wttrt,(Product=="Mechanical.wood.pulp"|Product=="Chemical.wood.pulp"|
                                 Product=="Pulp.from.fibres.other.than.wood"|Product=="Recovered.fibre.pulp"))#导入纸张净进口数据
CE_wttrt_Pulp_W=dcast(CE_wttrt_Pulp,Region+Year~Product,value.var = "Value")
CE_wttrt_Pulp_W$Value=rowSums(CE_wttrt_Pulp_W[,3:6])
CE_wttrt_Pulp_SUM=CE_wttrt_Pulp_W[,-(3:6)]

###造纸排放到空气
CE_wttrt_Paper=subset(CE_wttrt,(Product=="Newsprint"|Product=="Printing.and.writing.papers"|
                                  Product=="Household.and.sanitary.papers"|Product=="Packaging.paper.and.paperboard"|
                                  Product=="Other.paper.and.paperboard"))#导入数据
CE_wttrt_Paper_W=dcast(CE_wttrt_Paper,Region+Year~Product,value.var = "Value")
CE_wttrt_Paper_W$Value=rowSums(CE_wttrt_Paper_W[,3:7])
CE_wttrt_Paper_SUM=CE_wttrt_Paper_W[,-(3:7)]

###制浆排放到自然水体
CE_wttrt_ToNature_Pulp=subset(CE_wttrt_ToNature,(Product=="Mechanical.wood.pulp"|Product=="Chemical.wood.pulp"|
                                                   Product=="Pulp.from.fibres.other.than.wood"|Product=="Recovered.fibre.pulp"))#导入纸张净进口数据
CE_wttrt_ToNature_Pulp_W=dcast(CE_wttrt_ToNature_Pulp,Region+Year~Product,value.var = "Value")
CE_wttrt_ToNature_Pulp_W$Value=rowSums(CE_wttrt_ToNature_Pulp_W[,3:6])
CE_wttrt_ToNature_Pulp_SUM=CE_wttrt_ToNature_Pulp_W[,-(3:6)]

###造纸排放到自然水体
CE_wttrt_ToNature_Paper=subset(CE_wttrt_ToNature,(Product=="Newsprint"|Product=="Printing.and.writing.papers"|
                                                    Product=="Household.and.sanitary.papers"|Product=="Packaging.paper.and.paperboard"|
                                                    Product=="Other.paper.and.paperboard"))#导入纸张净进口数据
CE_wttrt_ToNature_Paper_W=dcast(CE_wttrt_ToNature_Paper,Region+Year~Product,value.var = "Value")
CE_wttrt_ToNature_Paper_W$Value=rowSums(CE_wttrt_ToNature_Paper_W[,3:7])
CE_wttrt_ToNature_Paper_SUM=CE_wttrt_ToNature_Paper_W[,-(3:7)]

COD_RP=COD_W[,1:2]
COD_RP$Value=COD_W$Recovered.fibre.pulp

########################################### 2. 其次，针对30个国家模拟MFA #############################################################################
for (n in 1:30){
  FAOSTAT=read.xlsx(paste("INPUT/PaperMFA/PaperMFA_",RG[n,2],".xlsx",sep = "",collapse = NULL),sheetName = "FAOSTAT")
  MBP_YR=read.xlsx(paste("INPUT/PaperMFA/PaperMFA_",RG[n,2],".xlsx",sep = "",collapse = NULL),sheetName = "MBP_YR")
  MBP_AM=read.xlsx("INPUT/PaperMFA/PaperMFA_China.xlsx",sheetName = "MBP_AM")#MBP_AM是物质平衡的输入参数,分配参数
  YR=MBP_YR[1:10,2:60]#YR是产出率表格的参数数据
  AM=MBP_AM[1:20,3:61]#AM是纸浆配比数据
  ##导入和废水处理相关的碳流动
  CE_wttrt_Pulp_t=dcast(CE_wttrt_Pulp,Region+Product~Year,Value.Var="Value")##尾缀“t"表示this，本脚本中专用的意思
  ##制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_CWP=subset(CE_wttrt_Pulp_t,Region==RG[n,2]&Product=="Chemical.wood.pulp")[,3:61]##化学木浆制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_NWP=subset(CE_wttrt_Pulp_t,Region==RG[n,2]&Product=="Pulp.from.fibres.other.than.wood")[,3:61]##化学草浆制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_MP=subset(CE_wttrt_Pulp_t,Region==RG[n,2]&Product=="Mechanical.wood.pulp")[,3:61]##机械制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_RP=subset(CE_wttrt_Pulp_t,Region==RG[n,2]&Product=="Recovered.fibre.pulp")[,3:61]##废纸浆制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_ToNature_Pulp_t=dcast(CE_wttrt_ToNature_Pulp,Region+Product~Year,Value.Var="Value")##尾缀“t"表示this，本脚本中专用的意思
  ##制浆排出废水COD流到自然水体或者存在于淤泥中的碳
  CE_wttrt_ToNature_CWP=subset(CE_wttrt_ToNature_Pulp_t,Region==RG[n,2]&Product=="Chemical.wood.pulp")[,3:61]
  CE_wttrt_ToNature_NWP=subset(CE_wttrt_ToNature_Pulp_t,Region==RG[n,2]&Product=="Pulp.from.fibres.other.than.wood")[,3:61]
  CE_wttrt_ToNature_MP=subset(CE_wttrt_ToNature_Pulp_t,Region==RG[n,2]&Product=="Mechanical.wood.pulp")[,3:61]
  CE_wttrt_ToNature_RP=subset(CE_wttrt_ToNature_Pulp_t,Region==RG[n,2]&Product=="Recovered.fibre.pulp")[,3:61]##这里的COD是用原来计算的废纸浆产量估算的，不能使用
  COD_RP_n=subset(COD_RP,Region==RG[n,2])
  COD_RP_n=dcast(COD_RP_n,Region~Year,value.var = "Value")
  COD_RP_n=COD_RP_n[,-1]
  ##最终产品消费量
  FB21=FAOSTAT[8,3:61]+FAOSTAT[20,3:61]#Newsprint to Use
  FB21[FB21<0]=0
  FB22=FAOSTAT[9,3:61]+FAOSTAT[21,3:61]#Printing and writing papers to Use
  FB22[FB22<0]=0
  FB23=FAOSTAT[10,3:61]+FAOSTAT[22,3:61]#Household and sanitary papers to Use
  FB23[FB23<0]=0
  FB24=FAOSTAT[11,3:61]+FAOSTAT[23,3:61]#Packaging papers to Use
  FB24[FB24<0]=0
  FB25=FAOSTAT[12,3:61]+FAOSTAT[24,3:61]#Other papers to Use
  FB25[FB25<0]=0
  TPC=FB21+FB22+FB23+FB24+FB25#Total paper consumption
  ##制浆过程流动
  FB5=FAOSTAT[2,3:61]+FAOSTAT[3,3:61]*0.5#Mechanical pulping to Mechanical pulp
  FB5[FB5<0]=0
  FB1=FB5/as.numeric(YR[1,])#Wood to Mechanical pulping
  FB1[FB1<0]=0
  FB2=(FAOSTAT[3,3:61]*0.5+FAOSTAT[4,3:61])/as.numeric(YR[2,])#Wood to Chemical pulping
  FB2[FB2<0]=0
  FB3=FAOSTAT[5,3:61]/as.numeric(YR[2,])#Other fibres to Chemical pulping
  FB3[FB3<0]=0
  FB4=FAOSTAT[7,3:61]+FAOSTAT[19,3:61]#Paper for recycling to Recycled pulping
  FB4[FB4<0]=0
  FB6=FB1-FB5#Mechanical pulping to Pulping waste
  FB6[FB6<0]=0
  FB7=FAOSTAT[4,3:61]+FAOSTAT[5,3:61]+FAOSTAT[3,3:61]*0.5#Chemical pulping to Chemical pulp
  FB7[FB7<0]=0
  FB8=FB2+FB3-FB7#Chemical pulping to Pulping waste
  FB8[FB8<0]=0
  FB9=(FB4*(1-(FB21*0.1+FB22*as.numeric(AM[8,])+FB23*0+FB24*0.1+FB25*0.1)/(FB21+FB22+FB23+FB24+FB25))+19.385/(1.8646*0.45*1000))/
    (1+COD_RP_n/(1.8646*0.45*1000))##FB9的计算公式是用设x的方式解出来的，所以看起来有点难理解
  FB9[FB9<0]=0
  FB10=FB4-FB9#Recycled pulping to pulping waste
  FB10[FB10<0]=0
  NER_RP=FB10/2-(FB9*COD_RP_n-19.385)/(1.8646*2*0.45*1000)##废纸制浆过程中产生的被非能源回收处理的废料（即，废料料非生物质部分的1/2）
  NER_RP[NER_RP<0]=0
  LF_RP=NER_RP##假设填埋和非能源回收的量一样（假设来自Stijn）
  LF_RP[LF_RP<0]=0
  ER_CP=FB8-(CE_wttrt_CWP+CE_wttrt_NWP)*12/(25*16*0.45*1000)-(CE_wttrt_ToNature_CWP+CE_wttrt_ToNature_NWP)*12/(44*0.45*1000)##化学浆制浆过程产生的用于能源回收的废料量
  ER_CP[ER_CP<0]=0
  ER_MP=FB6-(CE_wttrt_MP)*12/(25*16*0.45*1000)-(CE_wttrt_ToNature_MP)*12/(44*0.45*1000)##机械浆制浆过程产生的用于能源回收的废料量
  ER_MP[ER_MP<0]=0
  ##上面两个公式除以0.45，是因为假设其他生物质随有机碳与木材结构等比例流失
  FB32=ER_CP+ER_MP#Pulping waste to Energy recovery,已经去除了随COD流失的C
  FB32[FB32<0]=0
  FB33=NER_RP#Pulping waste to Non-energy recovery，全部为非生物质
  FB33[FB33<0]=0
  FB34=LF_RP#Pulping waste to Landfill，全部为非生物质
  FB34[FB34<0]=0
  ##进出口
  FB35=FAOSTAT[14,3:61]+FAOSTAT[15,3:61]*0.5#Net import to Mechanical pulp
  FB36=FAOSTAT[16,3:61]+FAOSTAT[17,3:61]+FAOSTAT[15,3:61]*0.5#Net import to Chemical pulp
  FB37=FAOSTAT[18,3:61]#Net import to Recycled pulp(Net import of Recovered pulp)
  FB38=FAOSTAT[20,3:61]#Net import to Newsprint
  FB39=FAOSTAT[21,3:61]#Net import to Printing and writing papers
  FB40=FAOSTAT[22,3:61]#Net import to Household and sanitary papers
  FB41=FAOSTAT[23,3:61]#Net import to Packaging papers
  FB42=FAOSTAT[24,3:61]#Net import to Other papers
  FB43=FAOSTAT[19,3:61]#Net import to Recycled pulping(Net import of Recovered paper)
  FB11=FB5+FB35#Mechanical pulp to Papermaking
  FB11[FB11<0]=0
  FB12=FB7+FB36#Chemical pulp to Papermaking
  FB12[FB12<0]=0
  FB13=FB9+FB37#Recycled pulp to Papermaking
  FB13[FB13<0]=0
  FB15=FAOSTAT[8,3:61]#Papermaking to Newsprint
  FB15[FB15<0]=0
  FB16=FAOSTAT[9,3:61]#Papermaking to Printing and writing papers
  FB16[FB16<0]=0
  FB17=FAOSTAT[10,3:61]#Papermaking to Household and sanitary papers
  FB17[FB17<0]=0
  FB18=FAOSTAT[11,3:61]#Papermaking to Packaging papers
  FB18[FB18<0]=0
  FB19=FAOSTAT[12,3:61]#Papermaking to Other papers
  FB19[FB19<0]=0
  FB14=(FB15*as.numeric(AM[4,])+FB16*as.numeric(AM[8,])+FB18*as.numeric(AM[16,])+FB19*as.numeric(AM[20,]))/as.numeric(YR[4,])#Non-fibrous to Papermaking
  FB14[FB14<0]=0
  FB20=(FB11+FB12+FB13+FB14)-(FB15+FB16+FB17+FB18+FB19)-
    CE_wttrt_Paper_SUM$Value[(59*n-58):(59*n)]*12/(25*16*0.45*1000)-
       CE_wttrt_ToNature_Paper_SUM$Value[(59*n-58):(59*n)]*12/(44*0.45*1000)#Papermaking to Paper for recycling
  FB20[FB20<0]=0
  FB26=TPC*as.numeric(YR[5,])#Use to Stock
  FB26[FB26<0]=0
  FB27=FAOSTAT[7,3:61]-FB20#Use to Paper for recycling【2023.1.19修改：之前FB27是包含FB20的，现在为了桑基图配平摘出来】
  FB27[FB27<0]=0
  FB28=(TPC-FB26-FB27)*as.numeric(YR[8,])#Use to Incineration【2023.1.21修改：摘出来FB27，那么这个公式就不能加FB20了】
  FB28[FB28<0]=0
  FB29=(TPC-FB26-FB27)*as.numeric(YR[6,])#Use to Energy recovery【2023.1.21修改：摘出来FB27，那么这个公式就不能加FB20了】
  FB29[FB29<0]=0
  FB30=(TPC-FB26-FB27)*as.numeric(YR[7,])#Use to Non-energy recovery【2023.1.21修改：摘出来FB27，那么这个公式就不能加FB20了】
  FB30[FB30<0]=0
  FB31=TPC-FB26-FB27-FB28-FB29-FB30#Use to Landfill
  FB31[FB31<0]=0
  FB44=FB6+FB8+FB10-(FB32+FB33+FB34)#【2023.1.19修改：之前计算了由于废水处理损失的部分，但是没有设置这个flow，现在加上了】
  FB44[FB44<0]=0
  FB45=FB11+FB12+FB13+FB14-(FB15+FB16+FB17+FB18+FB19+FB20)#【2023.1.20修改：之前计算了由于废水处理损失的部分，但是没有设置这个flow，现在加上了】
  FB45[FB45<0]=0

  FB=FB1
  for(f in 2:45){
    FB=rbind(FB,get(paste("FB",f, sep="")))
  }
  FB=cbind(FlowName[,],FB)
  write.xlsx(FB,file=paste("OUT/MFA/MF_",RG[n,2],".xlsx",sep = "",collapse = NULL),sheetName=RG[n,2])
}

################################## 3. 上述表格整合到一张表格里 ######################################################################
MF_1=read.xlsx(paste("OUT/MFA/MF_",RG[1,2],".xlsx",sep = "",collapse = NULL),sheetName=RG[1,2])##读入第一个国家的物质流数据
MF_1[,1]=RG[1,2]##去掉无用的序号列
MF=MF_1##给MF附上第一组值
##写循环整合成一张表
for (n in 2:30){
  MF_n=read.xlsx(paste("OUT/MFA/MF_",RG[n,2],".xlsx",sep = "",collapse = NULL),sheetName=RG[n,2])##读入第n个国家的物质流数据
  MF_n[,1]=RG[n,2]##将第一列改为国家
  MF=rbind(MF,MF_n)##将第n个国家的数据整合到数据框MF中去
}
MF[is.na(MF)]=0##将MF中的缺失值都替换成0
write.xlsx(MF,file="OUT/MFA/Flows_30Nations.xlsx",sheetName="Flows")##导出数据

############################### 4. 重要的手动操作 ###################################################################################
##首先，将Flows_30Nations.xlsx的格式稍作修改（预计耗时2分钟）：
##①首列删除
##②国家那一列加上列名"Region"
##③年份将X去掉，改为数字格式的1961-2019
##④俄罗斯1992年之前的数据都替换为0
##⑤将该表格复制到Global Paper_Info.xlsx中去
##⑥将表格中各国的FB9数据复制到Global Paper_Info.xlsx的FAOSTAT的Recycled pulp Production行去






