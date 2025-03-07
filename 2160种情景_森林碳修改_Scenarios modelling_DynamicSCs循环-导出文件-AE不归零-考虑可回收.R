#######ScenarioS simulation###########
##搜索“参数修改点”，直达情景设置位置
##2023年5月份修改：
##1. ScenarioModelling_Info中的ES2050表格里，biomass and waste一列的值去掉，只使用非生物质能源部分的比例，即将非生物质能源看作百分百，再分配。
##2. 回收率的分母部分改成了可回收的部分
##3. 因为避免的排放比较复杂（尤其是未来某些国家电力和热力可以负排放这个原因），原来方法的CE_ExBio可能在某些位点的计算有问题
##   所以新方法中在能源计算那里就判断该国该行业生物质供能与总能量需求的相对大小，当大于时，也不管多出来的生物质能避免多少排放了
##   （因为当能源系统的使用变成负排放时，不用反而比用还差，太tricky了，所以未来情景中不考虑这块了，这里也没必要算了）；当小于时，
##   那么生物质不能满足的部分就由外部能源供应提供，按部就班地计算排放就好了


##运行以下三行代码，保证内存
options(java.parameters='-Xms8g')
gc()###内存不够时跑一下这个
memory.limit(1024000)

##导入以下包
install.packages("rJava") 
install.packages("xlsxjars") 
install.packages("RJDBC")
install.packages("plyr")
install.packages("data.table")
install.packages("ggplot2")
install.packages("reshape2")
install.packages("do")
install.packages("stringr")
install.packages("cowplot")
install.packages("RColorBrewer")
install.packages("plyr")
install.packages("basicTrendline")
install.packages("xlsx") 
install.packages("openxlsx")
install.packages("XLConnect")
##引用以下包
library(rJava)
library(RJDBC)
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
library(xlsx)

##设置默认路径，方便后续数据读取和存储
setwd("D:/坚果云同步文件夹/RBOOK/Global Paper")

#############################常用数据，提前导入#####################################
Flows_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = 'Flows')
Flows=melt(Flows_W,id=c("Region","Flow","Input","Output"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
RG=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "Regions")

##################################### 1. 完整计算过程，循环输入参数即可完成对所有情景的模拟，首先生成参数表格 ################################

##构建所有情景中净排放表格
AllScs=RG[1:30,2:3]##构建基本表格，方便后面数据存入
##构建所有情景中净排放表格
AllScs_ExAE=RG[1:30,2:3]##构建基本表格，方便后面数据存入
##构建所有情景中净排放表格
AllScs_ExAEprd=RG[1:30,2:3]##构建基本表格，方便后面数据存入
##构建情景编号列表
hangName=c("a","b","c","d","e","h","l")##行名
Content=c("ES","DMD","RR","YR","SEC","SR","CR")##行名代表的含义
SC_Code=data.frame(hangName,Content)##构建数据框
for (m in 1:3){
  a=m
  f=m
  g=m
  k=m
  for (b in 1:1){
    for (c in 1:5){
      for (d in 1:4){
        for (e in 1:3){
          for (h in 1:4){
            for (l in 1:3){
              SC_Code_n=c(a,b,c,d,e,h,l)
              SC_Code=cbind(SC_Code,SC_Code_n)
              
              
            }
            
            
            
          }
          
        }
        
        
      }
      
    }
    
    
  }
}##编制情景编号列表
write.csv(SC_Code,"OUT/Scenario/All scenarios (2160)/SC_Code.csv")##存一下情景编号列表结果


##正式跑情景
for (pq in 1:2160){
  ##############读取列表编号##############
  a=SC_Code[1,pq+2]
  f=SC_Code[1,pq+2]
  g=SC_Code[1,pq+2]
  k=SC_Code[1,pq+2]
  b=SC_Code[2,pq+2]
  c=SC_Code[3,pq+2]
  d=SC_Code[4,pq+2]
  e=SC_Code[5,pq+2]
  h=SC_Code[6,pq+2]
  l=SC_Code[7,pq+2]
  ############################# 【参数修改点1】电力排放系数 Elc_CI: a ############################################
  ##导入电力排放系数
  Elc_CI_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="Elc_CI_2050")
  Elc_CI_ScName=unique(Elc_CI_Data$Scenario)
  Elc_CI=subset(Elc_CI_Data,Scenario==Elc_CI_ScName[a])##取出电力系数情景数据【情景修改点】
  
  ########################################## 一、物质流模拟 ###############################################
  ################################ 【参数修改点2】需求量 DMD: b ##############################################
  ##已经在excel里面计算了2050年的5种纸产品需求量，共3种情景，一起导入
  ##DMD_W表示demand_widedata
  DMD_W=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName ='DMD_2050_W')
  ##将宽数据转化为长数据，DMD_L表示demand_longdata
  DMD_L=melt(DMD_W,id=c("Region","Scenario"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
  ##【参数修改点】需求量
  ##DMD_S表示demand in scenarios
  DMD_ScName=unique(DMD_L$Scenario)##定义需求量情景的名称
  DMD_S=subset(DMD_L,Scenario==DMD_ScName[b])#选择消费量情景
  DMD_S=DMD_S[,-2]#去掉第二列Scenario
  
  
  
  ##读取纸产品净进口数据
  NIperCON_2019=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName ='NIperCON_2019')
  NIperCON_2019_Paper=subset(NIperCON_2019,Product=="NP"|Product=="PW"|Product=="HS"|Product=="PP"|Product=="OP")
  NIperCON_2019_Pulp=subset(NIperCON_2019,Product=="CWP"|Product=="MP"|Product=="RP"|Product=="NWP")
  ##Demand_Scenario_Widedata
  DMD_S_W=dcast(DMD_S,Region~Product,value.var = "Value")
  ##求和
  DMD_S_W$SUM=rowSums(DMD_S_W[,2:6])
  ##各种纸张消费量总量
  PaperCON=DMD_S_W#构造纸消费量数据框
  PaperCON$Value=DMD_S_W$SUM#纸张消费量这个数据框里就只有总量数据，DMD_S_W中有分产品的和总量
  PaperCON=PaperCON[,-(2:7)]#去掉原来DMD_S_W分产品消费量和总消费量
  ##纸张消费量中的生物质如下
  PaperCON_Bio=DMD_S_W
  PaperCON_Bio$Value=DMD_S_W$NP*0.9+DMD_S_W$PW*0.7+DMD_S_W$HS*1+DMD_S_W$PP*0.9+DMD_S_W$OP*0.9
  PaperCON_Bio=PaperCON_Bio[,-(2:7)]
  
  ##################################### 1. 将纸张消费量分配到各产品 ####################################################
  SF21=DMD_S_W
  SF21$Value=DMD_S_W$NP
  SF21=SF21[,-(2:7)]
  SF22=SF21
  SF22$Value=DMD_S_W$PW
  SF23=SF21
  SF23$Value=DMD_S_W$HS
  SF24=SF21
  SF24$Value=DMD_S_W$PP
  SF25=SF21
  SF25$Value=DMD_S_W$OP
  
  ##################################### 2. 纸张生产量 ####################################################################
  DMD_S=DMD_S[order(DMD_S$Region),]
  PaperPRD=DMD_S
  PaperPRD$Value=DMD_S$Value*(1-as.numeric(NIperCON_2019_Paper$Value))
  ##分配到各产品
  SF15=SF21
  SF15$Value=subset(PaperPRD,Product=="NP")$Value
  SF16=SF21
  SF16$Value=subset(PaperPRD,Product=="PW")$Value
  SF17=SF21
  SF17$Value=subset(PaperPRD,Product=="HS")$Value
  SF18=SF21
  SF18$Value=subset(PaperPRD,Product=="PP")$Value
  SF19=SF21
  SF19$Value=subset(PaperPRD,Product=="OP")$Value
  
  ###################################### 3.消费量去向 ####################################################################
  ##先把SF20计算出来
  SF20=SF21
  SF20$Value=(SF15$Value+SF16$Value+SF17$Value+SF18$Value+SF19$Value)*0.05/0.95
  ##去除非生物质，用于后续部分碳排放计算
  SF20_Bio=SF20
  SF20_Bio$Value=(SF15$Value*0.9+SF16$Value*0.7+SF17$Value*1+SF18$Value*0.9+SF19$Value*0.9)*0.05/0.95
  
  ###################################### 3.1 产品存量 #####################################################################
  SF26_Bio=PaperCON_Bio
  SF26_Bio$Value=PaperCON_Bio$Value*0.09
  SF26=SF26_Bio
  SF26$Value=PaperCON$Value*0.09
  
  ######################################### 【参数修改点3】回收比例 RR: c ################################################
  ###################################### 3.2 回收量 #######################################################################
  ##回收利用比例
  ##导入回收占消费后废物的比例（消费量扣除存量）
  RcyclRate2050_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "RcyclRate2050_exHS")
  RcyclRate2050_ScName=unique(RcyclRate2050_Data$Scenario)
  RcyclRate2050=subset(RcyclRate2050_Data,Scenario==RcyclRate2050_ScName[c])
  ##计算回收量
  SF27=SF21
  SF27$Value=(PaperCON$Value*0.91+SF20$Value-SF23$Value)*RcyclRate2050$Value
  SF23_Bio=SF21
  SF23_Bio$Value=SF23$Value*1
  SF27_Bio=SF21
  SF27_Bio$Value=(PaperCON_Bio$Value*0.91+SF20_Bio$Value-SF23_Bio$Value)*RcyclRate2050$Value
  ###################################### 3.3 焚烧量 #######################################################################
  ###################################### 【参数修改点4】垃圾处理比例 YR: d ######################################################
  ##【重要】参数设置：废物处理比例需要替换多个
  YR2050_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName ='YR2050')#导入生活垃圾处理比例数据
  YR_ScName=YR2050_Data$Scenario
  YR_ScName=unique(YR_ScName)
  YR2050=subset(YR2050_Data,Scenario==YR_ScName[d])
  SF28=SF27
  SF28$Value=(PaperCON$Value+SF20$Value-SF26$Value-SF27$Value)*YR2050$PCW_inc
  SF28_Bio=SF28
  SF28_Bio$Value=(PaperCON_Bio$Value+SF20_Bio$Value-SF26_Bio$Value-SF27_Bio$Value)*YR2050$PCW_inc
  
  ###################################### 3.4 能源回收量 #######################################################################
  SF29=SF28
  SF29$Value=(PaperCON$Value+SF20$Value-SF26$Value-SF27$Value)*YR2050$PCW_er
  SF29_Bio=SF29
  SF29_Bio$Value=(PaperCON_Bio$Value+SF20_Bio$Value-SF26_Bio$Value-SF27_Bio$Value)*YR2050$PCW_er
  
  ###################################### 3.5 非能源回收量 #######################################################################
  SF30=SF29
  SF30$Value=(PaperCON$Value+SF20$Value-SF26$Value-SF27$Value)*YR2050$PCW_ner
  SF30_Bio=SF30
  SF30_Bio$Value=(PaperCON_Bio$Value+SF20_Bio$Value-SF26_Bio$Value-SF27_Bio$Value)*YR2050$PCW_ner
  
  ###################################### 3.6 填埋量 #######################################################################
  SF31=SF30
  SF31$Value=(PaperCON$Value+SF20$Value-SF26$Value-SF27$Value)-SF28$Value-SF29$Value-SF30$Value
  SF31_Bio=SF31
  SF31_Bio$Value=(PaperCON_Bio$Value+SF20_Bio$Value-SF26_Bio$Value-SF27_Bio$Value)-SF28_Bio$Value-SF29_Bio$Value-SF30_Bio$Value
  
  ##################################### 4. 计算废纸消费量#################################
  ##读取废纸净进口占消费量的比例
  NIperPRD_2019_RcycPaper=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="NIperPRD_2019_RcycPaper")#Excel里先构造个表格，计算净进口回收纸张占消费量比例，然后读取数据
  NIperPRD_2019_RcycPaper=NIperPRD_2019_RcycPaper[1:30,1:4]
  ##计算废纸消费量
  SF4=SF21
  SF4$Value=(SF20$Value+SF27$Value)*(1+as.numeric(NIperPRD_2019_RcycPaper$Value))
  
  ##计算其中的生物质
  SF4_Bio=SF4
  SF4_Bio$Value=SF4$Value*(1-(SF21$Value*0.1+SF22$Value*0.3+SF23$Value*0+SF24$Value*0.1+SF25$Value*0.1)/(SF21$Value+SF22$Value+SF23$Value+SF24$Value+SF25$Value))
  ##计算废纸浆生产量
  ##首先，要计算废纸浆生产过程中的COD产生量
  COD_W=read.xlsx("INPUT/Global Paper_Info.xlsx",sheetName = "WtTrt_COD")
  COD_W_2019=subset(COD_W,Year==2019&Region!="World")
  COD_RP_2019=COD_W_2019[,-(2:10)]
  names(COD_RP_2019)=c("Region","Value")
  ##然后通过解方程得出SF9，SF9指的是废纸浆生产量
  SF9=SF4
  SF9$Value=(SF4_Bio$Value+19.385/(1.8646*0.45*1000))/
    (1+COD_RP_2019$Value/(1.8646*0.45*1000))##SF9的计算公式是用设x的方式解出来的，所以看起来有点难理解
  ##计算废纸浆生产过程中的工业垃圾产生量
  SF10=SF4
  SF10$Value=SF4$Value-SF9$Value
  
  ##################################### 5. 纸浆消费量（也包括填料和涂料）####################################################
  PulpCON=PaperPRD
  PulpCON$Value=PaperPRD$Value/0.95#加上造纸过程的损失量，得到总投入量
  PulpCON_W=dcast(PulpCON,Region~Product,value.var = "Value")
  PulpCON_Bio=PulpCON_W
  PulpCON_Bio$Value=PulpCON_W$NP*0.9+PulpCON_W$PW*0.7+PulpCON_W$HS*1+PulpCON_W$PP*0.9+PulpCON_W$OP*0.9#每种纸的重量乘以其中生物质含量比例
  PulpCON_Bio=PulpCON_Bio[,-(2:6)]
  ##计算非纤维材料消费量
  SF14=SF21
  SF14$Value=rowSums(PulpCON_W[,2:6])-PulpCON_Bio$Value
  
  ##################################### 6. 纸浆分结构消费量#####################################
  ##计算废纸浆消费量
  SF13=SF21
  ##读取废纸浆净进口占生产量比例
  NIperPRD_2019_RcycPulp=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="NIperPRD_2019_RcycPulp")#Excel里先构造个表格，计算净进口回收纸张占产量比例，然后读取数据
  ##计算废纸浆消费量
  SF13=SF21
  SF13$Value=SF9$Value*(1+NIperPRD_2019_RcycPulp$Value)
  PulpCON_ExRP=SF21
  for (n in 1:30){
    if (SF13$Value[n]>PulpCON_Bio$Value[n]*0.9){
      SF13$Value[n]=PulpCON_Bio$Value[n]*0.9
      PulpCON_ExRP$Value[n]=PulpCON_Bio$Value[n]*0.1
    }else{
      PulpCON_ExRP$Value[n]=rowSums(PulpCON_W[n,2:6])-SF14$Value[n]-SF13$Value[n]
    }
  }
  
  ##将剩余纸浆消费量分配给机械木浆、化学木浆、草浆
  ##读取2019年机械木浆、化学木浆、草浆分别占总消费量的比例
  PulpStructure_2019=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName ='PulpStructure_2019')#这里是2019年的纸浆结构
  ##计算机械木浆消费量
  SF11=SF21
  SF11$Value=PulpCON_ExRP$Value*PulpStructure_2019$MP.C_percent
  ##计算化学木浆消费量
  SF12_CWP=SF21
  SF12_CWP$Value=PulpCON_ExRP$Value*PulpStructure_2019$CWP.C_percent
  ##计算草浆消费量
  SF12_NWP=SF21
  SF12_NWP$Value=PulpCON_ExRP$Value*PulpStructure_2019$NWP.C_percent
  
  ##################################### 7. 纸浆产量################################################
  ##机械制浆产量
  SF5=SF21
  SF5$Value=SF11$Value*(1-as.numeric(subset(NIperCON_2019,Product=="MP")$Value))
  ##机械纸浆废物
  SF6=SF21
  SF6$Value=SF5$Value*0.08/0.92
  ##化学木浆产量
  SF7_CWP=SF21
  SF7_CWP$Value=SF12_CWP$Value*(1-as.numeric(subset(NIperCON_2019,Product=="CWP")$Value))
  ##草浆产量
  SF7_NWP=SF21
  SF7_NWP$Value=SF12_NWP$Value*(1-as.numeric(subset(NIperCON_2019,Product=="NWP")$Value))
  ##化学木浆废物
  SF8_CWP=SF21
  SF8_CWP$Value=SF7_CWP$Value*0.53/0.47
  ##草浆废物
  SF8_NWP=SF21
  SF8_NWP$Value=SF7_NWP$Value*0.53/0.47
  
  ##################################### 8. 构造纸浆和纸张产量数据框#########################################
  ProductPRD=cbind(SF7_CWP,SF17[,2],SF5[,2],SF15[,2],SF7_NWP[,2],SF19[,2],SF18[,2],SF16[,2],SF9[,2])
  names(ProductPRD)=c("Region","CWP","HS","MP","NP","NWP","OP","PP","PW","RP")
  ProductPRD_L=melt(ProductPRD,id=c("Region"),variable.name = "Product",value.name = "Value")
  
  ##这部分涉及COD比较复杂
  ##先计算如下部分
  ####################################### 9. 废水处理过程碳排放 #########################################################
  COD=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "Wttrt_COD_2019")#导入单位产品的COD产生量
  COD_L=melt(COD,id=c("Region"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
  Removal=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "Wttrt_Removal_2019")#导入COD去除量
  Removal_L=melt(Removal,id=c("Region"),variable.name = "Product",value.name = "Value")
  MCF=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "Wttrt_MCF_2019")#导入甲烷修正因子
  MCF_L=melt(MCF,id=c("Region"),variable.name = "Product",value.name = "Value")
  CE_wttrt=COD_L#构建基本的数据框
  ProductPRD_L_ExPR=subset(ProductPRD_L,Product!="PR")#把印刷数据去掉
  CE_wttrt$Value=ProductPRD_L_ExPR$Value*COD_L$Value*Removal_L$Value*MCF_L$Value*0.25*25#乘以25转化为CO2当量，单位是kg
  
  
  ####################################### 9.1 排向自然界的碳 ##########################################
  CE_wttrt_ToNature=COD_L[,1:2]
  CE_wttrt_ToNature$Value=((ProductPRD_L_ExPR$Value*COD_L$Value-19.385)*(1/1.8646)-CE_wttrt$Value*12/(25*16))*44/12#单位是kg CO2
  CE_wttrt_ToNature[CE_wttrt_ToNature<0]=0##该方法失灵
  ##所以只能用以下条件判断语句实现该功能了
  for (n in 1:30){
    if(CE_wttrt_ToNature$Value[n]<0){
      CE_wttrt_ToNature$Value[n]=0
    }else{
      CE_wttrt_ToNature$Value[n]=CE_wttrt_ToNature$Value[n]
    }
  }##成功！
  
  
  
  ####################################### 9.2 制浆排放到空气 ############################################
  CE_wttrt_Pulp=subset(CE_wttrt,(Product=="MP"|Product=="CWP"|
                                   Product=="NWP"|Product=="RP"))
  CE_wttrt_Pulp_W=dcast(CE_wttrt_Pulp,Region~Product,value.var = "Value")
  CE_wttrt_Pulp_W$Value=rowSums(CE_wttrt_Pulp_W[,2:5])
  CE_wttrt_Pulp_SUM=CE_wttrt_Pulp_W[,-(2:5)]
  
  ####################################### 9.3 造纸排放到空气 ###########################################
  CE_wttrt_Paper=subset(CE_wttrt,(Product=="NP"|Product=="PW"|
                                    Product=="HS"|Product=="PP"|
                                    Product=="OP"))
  CE_wttrt_Paper_W=dcast(CE_wttrt_Paper,Region~Product,value.var = "Value")
  CE_wttrt_Paper_W$Value=rowSums(CE_wttrt_Paper_W[,2:6])
  CE_wttrt_Paper_SUM=CE_wttrt_Paper_W[,-(2:6)]
  
  
  ####################################### 9.4 制浆排放到自然水体 ###########################################
  CE_wttrt_ToNature_Pulp=subset(CE_wttrt_ToNature,(Product=="MP"|Product=="CWP"|
                                                     Product=="NWP"|Product=="RP"))
  CE_wttrt_ToNature_Pulp_W=dcast(CE_wttrt_ToNature_Pulp,Region~Product,value.var = "Value")
  CE_wttrt_ToNature_Pulp_W$Value=rowSums(CE_wttrt_ToNature_Pulp_W[,2:5])
  CE_wttrt_ToNature_Pulp_SUM=CE_wttrt_ToNature_Pulp_W[,-(2:5)]
  
  ####################################### 9.5 造纸排放到自然水体 ###########################################
  CE_wttrt_ToNature_Paper=subset(CE_wttrt_ToNature,(Product=="NP"|Product=="PW"|
                                                      Product=="HS"|Product=="PP"|
                                                      Product=="OP"))
  CE_wttrt_ToNature_Paper_W=dcast(CE_wttrt_ToNature_Paper,Region~Product,value.var = "Value")
  CE_wttrt_ToNature_Paper_W$Value=rowSums(CE_wttrt_ToNature_Paper_W[,2:6])
  CE_wttrt_ToNature_Paper_SUM=CE_wttrt_ToNature_Paper_W[,-(2:6)]
  
  
  ##################################### 10. 原料消费量#########################################################
  ##废纸浆原料消费量：其实之前已经计算过了，这里重复计算，而且有问题，幸好后面没再用到该数据，对碳排放计算没有产生影响
  #SF4=SF21
  #SF4$Value=SF9$Value/0.75
  
  ##草浆原料消费量（假设均为本地供应，无进出口,因此草浆原料产量等于消费量）
  SF3=SF21
  SF3$Value=SF7_NWP$Value+SF8_NWP$Value
  ##化学浆原料木材消费量
  SF2=SF21
  SF2$Value=SF7_CWP$Value+SF8_CWP$Value
  ##机械浆原料木材消费量
  SF1=SF21
  SF1$Value=SF5$Value+SF6$Value
  
  ##################################### 11. 原料产量#########################################################
  WoodCON=SF21
  WoodCON$Value=(SF1$Value+SF2$Value)/(1-0.056)####加上损失的部分，推导出总输入木材量
  ##读入历史的净进口比例
  ##这里有个问题，那就是1990-2019年浆木的进出口数据缺失，共想出以下3个解决方案：
  ##①利用先前计算的1961-2019年的木浆生产原料消费量作为浆木消费量，然后结合浆木产量，计算得到各国浆木净进口量，然后计算净进口占消费量比例
  ##但是该方法得到的结果与某些相关统计数据以及常识直觉相矛盾（比如加拿大需要大量进口浆木）
  ##②利用1989年及之前的净进口比例代替2019年的缺失数据，但当时全球贸易远不如现在发达，所以使用该方法可能会引起较大争议
  ##③直接使用浆木上一级的工业圆木的净进口占消费量比例来代替浆木的净进口比例，该方法也存在一定的问题，但是是目前最好的方法【采用】
  NIperCON_2019_Wood=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="NIperCON_2019_Wood")#Excel里先构造个表格，计算净进口浆木占消费量比例，然后读取数据
  ##计算浆木产量，注意这里是重量，而不是体积
  WoodPRD=SF21
  WoodPRD$Value=WoodCON$Value*(1-NIperCON_2019_Wood$Value)
  WoodPRD$Value=WoodPRD$Value/0.54##将木材重量转化成木材体积。0.54ton/m3是木材的密度，这里的估算使用所有木材的平均密度，不分区、不分树种。
  ##################################### 12. 工业废物#########################################################
  ##导入和废水处理相关的碳流动
  ##制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_CWP=subset(CE_wttrt_Pulp,Product=="CWP")[,3]##化学木浆制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_NWP=subset(CE_wttrt_Pulp,Product=="NWP")[,3]##化学草浆制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_MP=subset(CE_wttrt_Pulp,Product=="MP")[,3]##机械制浆排出废水COD去除过程产生的CH4的25倍
  CE_wttrt_RP=subset(CE_wttrt_Pulp,Product=="RP")[,3]##废纸浆制浆排出废水COD去除过程产生的CH4的25倍
  ##制浆排出废水COD流到自然水体或者存在于淤泥中的碳
  CE_wttrt_ToNature_CWP=subset(CE_wttrt_ToNature_Pulp,Product=="CWP")[,3]
  CE_wttrt_ToNature_NWP=subset(CE_wttrt_ToNature_Pulp,Product=="NWP")[,3]
  CE_wttrt_ToNature_MP=subset(CE_wttrt_ToNature_Pulp,Product=="MP")[,3]
  CE_wttrt_ToNature_RP=subset(CE_wttrt_ToNature_Pulp,Product=="RP")[,3]
  
  NER_RP=(SF10$Value-(CE_wttrt_RP)*12/(25*16*0.45*1000)-(CE_wttrt_ToNature_RP)*12/(44*0.45*1000))/2##废纸制浆过程中产生的被非能源回收处理的废料（即，废料料非生物质部分的1/2）
  NER_RP[NER_RP<0]=0
  LF_RP=NER_RP##假设填埋和非能源回收的量一样（假设来自Stijn）
  LF_RP[LF_RP<0]=0
  ER_CP=(SF8_CWP$Value+SF8_NWP$Value)-(CE_wttrt_CWP+CE_wttrt_NWP)*12/(25*16*0.45*1000)-(CE_wttrt_ToNature_CWP+CE_wttrt_ToNature_NWP)*12/(44*0.45*1000)##化学浆制浆过程产生的用于能源回收的废料量
  ER_CP[ER_CP<0]=0
  ER_MP=SF6$Value-(CE_wttrt_MP)*12/(25*16*0.45*1000)-(CE_wttrt_ToNature_MP)*12/(44*0.45*1000)##机械浆制浆过程产生的用于能源回收的废料量
  ER_MP[ER_MP<0]=0
  
  SF32=SF21
  SF32$Value=ER_CP+ER_MP
  ##该部分全部为生物质，因此不用专门计算生物质部分
  SF33=SF32
  SF33$Value=NER_RP
  ##SF34完全等同于SF33
  SF34=SF33
  
  
  ######################################## 二、能源使用和碳排放 #########################################################
  ########################################（一）上游的碳排放 #####################################################
  ####################################### 1.原料收集阶段能耗及碳排放 #####################################################
  ####################################### 1.1木材收集过程 ################################################################
  ####################################### 1.1.1木材收集过程能耗 ##########################################################
  E_wh=WoodPRD
  E_wh$Value=WoodPRD$Value*0.54*(0.13+0.55)#木材收集过程能耗，单位GJ。这里的木材使用量要用重量
  E_wh[E_wh<0]=0
  
  ####################################### 1.1.2 木材收集过程碳排放 ##############################################
  CE_wh=WoodPRD
  CE_wh$Value=WoodPRD$Value*0.54*0.13*74066.7/1000+WoodPRD$Value*0.54*0.55*Elc_CI$CI/1000##0.54ton/m3是木材的密度，这里的估算使用所有木材的平均密度，不分区、不分树种。
  CE_wh[CE_wh<0]=0
  
  ####################################### 1.2 非木材收集过程###############################################
  ####################################### 1.2.1 非木材收集能耗#############################################
  E_nwh=E_wh
  E_nwh$Value=(SF3$Value)*0.3142#非木材原料收集能耗，单位GJ
  E_nwh[E_nwh<0]=0
  
  ####################################### 1.2.2 非木材收集碳排放#############################################
  CE_nwh=E_wh
  CE_nwh$Value=(SF3$Value)*23.3#非木材原料收集碳排，单位kg
  CE_nwh[CE_nwh<0]=0
  
  ####################################### 1.3废纸收集过程 ##############################################################################
  ####################################### 1.3.1废纸收集能耗#######
  E_rpc=E_wh
  E_rpc$Value=SF27$Value*(0.0314+0.0702)#废纸收集能耗,GJ
  E_rpc[E_rpc<0]=0
  
  ####################################### 1.3.2废纸收集碳排放########
  CE_rpc=E_wh
  CE_rpc$Value=SF27$Value*0.0314*74.0667+SF27$Value*0.0702*Elc_CI$CI/1000#废纸收集过程碳排放，单位kg
  CE_rpc[CE_rpc<0]=0
  
  ####################################### 1.4原料收集过程######
  ####################################### 1.4.1原料收集总能耗#######
  E_he=E_wh
  E_he$Value=E_wh$Value+E_nwh$Value+E_rpc$Value
  E_he[E_he<0]=0
  
  ##1.4.3原料收集总碳排#######
  CE_he=CE_wh
  CE_he$Value=as.numeric(CE_wh$Value)+CE_nwh$Value+as.numeric(CE_rpc$Value)#原料收集总共的碳排放
  CE_he[CE_he<0]=0
  
  ####################################### 2. 化学品生产能耗及碳排放##########################
  ####################################### 2.1化学品生产能耗 ##############################################
  Chem=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "Chem")
  Chem=Chem[order(Chem$Product),]
  E_chem=ProductPRD_L
  E_chem$Value=ProductPRD_L$Value*Chem$SEC#各种纸浆和纸张药品投入的能源消耗，GJ
  
  ####################################### 2.2化学品生产碳排放 ############################################
  ####################################### 【参数修改点9】化学品生产碳排强度 ChemCI: k #################################
  ChemCI_alpha=c(0.25,0.75,1)##分别对应B2DS,RTS,和Y2019
  CE_chem=ProductPRD_L
  CE_chem$Value=ProductPRD_L$Value*Chem$CI*ChemCI_alpha[k]#各种纸浆和纸张药品投入的碳排放,kg
  
  
  ####################################### （二）废物处理阶段的碳汇或碳排 #######################################################################
  ####################################### 1.消费品形成产品存量——碳汇（选用CS_ps_Bio）####################################################
  ########产品存量碳汇分产品计算##############
  ########排除了非生物质部分###########
  ##########【重要】参数设置：需求量###########
  ##########纸张中的碳汇############
  CS_ps_Bio=PaperCON_Bio
  CS_ps_Bio$Value=PaperCON_Bio$Value*0.09*1.65*1000#单位为kg CO2。这里的CS_ps_Bio对应的就是F26_Bio
  
  
  ####################################### 2.生活废物回收库存碳汇（不作计算）################################################################
  #相关计算方法没有完全确定,不过量不大，不太影响结论
  ########这一块和历史算法不同，这里只是计算了一下有多少物质存储在回收纸张中。不过反正不计入最终结果，也没影响。
  ########不过这块对于后续市政垃圾处理量有影响，因为后续的垃圾处理比例是要乘以总消费量减去存量再减去回收量的值，才能计算出各处理方式的物质量
  ##前面给出了废纸回收数据
  
  
  ####################################### 3.生活废物燃烧产生的碳排放（选用CE_inc_Bio）################################################################
  ################排除了非生物质###########
  #####【重要】参数设置：废物处理比例##########
  CE_inc_Bio=PaperCON_Bio
  CE_inc_Bio$Value=SF28_Bio$Value*0.45*1000*44/12#最终单位为kg.假设纸张的热值与木材一致
  
  ####################################### 4.生活废物能源回收 #################################################################################
  ####################################### 4.1 生活废物能源回收部分的碳排放（选用CE_er_mw_Bio）#########################
  ###########排除了非生物质############
  CE_er_mw_Bio=PaperCON_Bio
  CE_er_mw_Bio$Value=SF29_Bio$Value*0.45*1000*44/12#最终单位为kg
  
  
  ####################################### 4.2 生活废物回收的能源（选用E_er_mw）#########################
  ###########不用排除非生物质############
  E_er_mw=PaperCON_Bio
  E_er_mw$Value=SF29$Value*13#最终单位GJ
  
  
  ####################################### 4.3 能源回收避免的碳排放——生活废物（选用AE_er_mw）#########################
  AE_er_mw=PaperCON_Bio
  AE_er_mw$Value=SF29$Value*0.013*0.25*Elc_CI$CI#最终单位为kg CO2。【重要】参数设置——废物处理比例、电力排放系数。
  
  
  ####################################### 5.生活废物填埋###################################################################
  ##计算填埋的生物质量
  Flows_Bio_His=read.xlsx("OUT/Flows.xlsx",sheetName = "Flows_Bio")
  Flows_Bio_His=Flows_Bio_His[,-1]
  Landfill_Bio_His=subset(Flows_Bio_His,Flow=="F31_Bio")
  Landfill_Bio_2019=Landfill_Bio_His[,-(2:62)]##选出2019年
  names(Landfill_Bio_2019)=c("Region","Value")##重新给列命名
  Landfill_Bio=Landfill_Bio_2019
  Landfill_Bio$Value=SF31_Bio$Value*(1-YR2050$PCW_er-YR2050$PCW_ner-YR2050$PCW_inc)
  
  ##针对上述的2019年和2050年填埋量（生物质）进行插值
  Landfill_intp=matrix(0,32,1)
  Landfill_intp=data.frame(Landfill_intp)
  names(Landfill_intp)=c("Year")
  Landfill_intp$Year=2019:2050
  for (n in 1:30){
    Landfill_intp_n=approx(c(2019,2050),c(Landfill_Bio_2019$Value[n],Landfill_Bio$Value[n]),n=32)
    Landfill_intp_n=data.frame(Landfill_intp_n)
    names(Landfill_intp_n)=c("Year",RG[n,2])
    Landfill_intp_n=Landfill_intp_n[,-1]
    Landfill_intp_n=data.frame(Landfill_intp_n)
    names(Landfill_intp_n)=c(RG[n,2])
    Landfill_intp=cbind(Landfill_intp,Landfill_intp_n)
  }
  Landfill_intp_L=melt(Landfill_intp,id=c("Year"),variable.name = "Region",value.name = "Value")#将原始宽型数据重组为长型数据
  Flows_Bio=read.xlsx("OUT/Flows.xlsx",sheetName = "Flows_Bio")
  Flows_Bio=Flows_Bio[,-65]
  F31_Bio=subset(Flows_Bio,Flow=="F31_Bio")
  F31_Bio=F31_Bio[,-(1)]#去掉冗余的列
  F31_Bio=F31_Bio[,-(2:4)]#去掉冗余列
  F31_Bio_L=melt(F31_Bio,id=c("Region"),variable.name = "Year",value.name = "Value")#将原始宽型数据重组为长型数据
  F31_Bio_L=F31_Bio_L[order(F31_Bio_L$Region),]#排序
  F31_Bio_L$Year=1961:2019#给年份重新赋值
  F31_Bio_L=subset(F31_Bio_L,Year!=2019)#去掉重复年份
  cols=colnames(Landfill_intp_L)#获得Landfill_intp_L列名
  new_cols=c(cols[2],cols[1],cols[3])#构造新的顺序
  Landfill_intp_L=Landfill_intp_L[,new_cols]#调整列顺序，与F31_Bio一致
  Landfill_6150=rbind(F31_Bio_L,Landfill_intp_L)#历史填埋的生物质量与未来模拟量合并，得到1961-2050年的填埋生物质物质量
  ##调整一下顺序
  Landfill_6150=Landfill_6150[order(Landfill_6150$Year),]
  Landfill_6150=Landfill_6150[order(Landfill_6150$Region),]
  Landfill_6150[is.na(Landfill_6150)]=0
  ####################################### 5.0 生活废物填埋释放CO2（选用CO2_lf_mw_Bio）#############################################
  ##5.0和5.1比较特殊，需要历史数据，因此与历史计算相同的变量名包括了历史和未来数据，而其他项，则指的是2050年数据。5.0和5.1两项，2050年的专门标注出来了。
  CO2_lf_mw_Bio=matrix(0,2700,3)
  CO2_lf_mw_Bio_n=matrix(0,90,1)
  for (n in 1:30){
    Landfill_P=Landfill_6150$Value[(90*n-89):(90*n)]
    t=seq(1961,2050,1)
    for (i in 2:90) {
      j=1:i
      CO2_lf_mw_Bio_n[i,1]=sum((1-exp(-0.05))*Landfill_P[j]*0.7*0.45*0.5*0.5*
                                 (44/12)*exp(-0.05*(t[i]-t[j])))*1000#单位为kg
    }
    CO2_lf_mw_Bio[(90*n-89):(90*n),1]=RG[n,2]
    CO2_lf_mw_Bio[(90*n-89):(90*n),2]=1961:2050
    CO2_lf_mw_Bio[(90*n-89):(90*n),3]=CO2_lf_mw_Bio_n
  }
  CO2_lf_mw_Bio=data.frame(CO2_lf_mw_Bio)#矩阵转化成数据框才能加列名
  names(CO2_lf_mw_Bio)=c("Region","Year","Value")#添加列名
  
  ##专门取出2050年数据
  CO2_lf_mw_Bio2050=subset(CO2_lf_mw_Bio,Year==2050)#取出2050年数据
  CO2_lf_mw_Bio2050=CO2_lf_mw_Bio2050[,-2]#去掉没用的列
  
  ####################################### 5.1 生活废物填埋释放CH4（选用CE_lf_mw_Bio）#############################################
  ###########排除了非生物质############
  ######################################### 【参数修改点10】甲烷捕集率CR: l ############################################
  CR_1961to2050=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "CR_1961to2050")
  CR_Name=unique(CR_1961to2050$Scenario)
  CR=subset(CR_1961to2050,Scenario==CR_Name[l])
  CE_lf_mw_Bio=matrix(0,2700,3)
  CE_lf_mw_Bio_n=matrix(0,90,1)
  for (n in 1:30){
    Landfill_P=Landfill_6150$Value[(90*n-89):(90*n)]
    t=seq(1961,2050,1)
    for (i in 2:90) {
      j=1:i
      CE_lf_mw_Bio_n[i,1]=sum((1-exp(-0.05))*Landfill_P[j]*0.7*0.45*0.5*0.5*
                                (16/12)*exp(-0.05*(t[i]-t[j]))*
                                (1-CR$CR[i])*(1-0.05))*25*1000#单位为kg
    }
    CE_lf_mw_Bio[(90*n-89):(90*n),1]=RG[n,2]
    CE_lf_mw_Bio[(90*n-89):(90*n),2]=1961:2050
    CE_lf_mw_Bio[(90*n-89):(90*n),3]=CE_lf_mw_Bio_n
  }
  CE_lf_mw_Bio=data.frame(CE_lf_mw_Bio)#矩阵转化成数据框才能加列名
  names(CE_lf_mw_Bio)=c("Region","Year","Value")#添加列名
  
  ##专门取出2050年数据
  CE_lf_mw_Bio2050=subset(CE_lf_mw_Bio,Year==2050)#取出2050年数据
  CE_lf_mw_Bio2050=CE_lf_mw_Bio2050[,-2]#去掉没用的列
  
  
  ####################################### 5.2生活废物填埋产生CH4的收集#########################
  ####################################### 5.2.1生活废物填埋产生CH4收集后燃烧碳排放（选用CECP_lf_mw_Bio）#########################
  ########排除了非生物质###########
  CECP_lf_mw_Bio=CE_lf_mw_Bio2050[,1:2]
  CECP_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio2050$Value)*CR$CR[90]*44/(25*(1-CR$CR[90])*(1-0.05)*16)#最终单位为kg CH4
  CECP_lf_mw_Bio[is.na(CECP_lf_mw_Bio)]=0
  
  ####################################### 5.2.2生活废物填埋CH4收集供应的能量：GJ（选用E_lf_mw_Bio）#########################
  #####排除了非生物质########
  E_lf_mw_Bio=CECP_lf_mw_Bio[1:2]
  E_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio2050$Value)*CR$CR[90]/(25*(1-CR$CR[90])*(1-0.05))*0.048*1000*0.35#产生的甲烷单位为kg, 天然气热值为0.048TJ/ton，假设该部分能量要转换成最后结果单位为GJ
  E_lf_mw_Bio[is.na(E_lf_mw_Bio)]=0
  ####################################### 5.2.3生活废物填埋产生的CH4收集量避免的排放（选用AE_lf_mw_Bio）#########################
  ########排除了非生物质###########
  ##【重要参数：电网排放系数】
  AE_lf_mw_Bio=CE_lf_mw_Bio2050[1:2]
  AE_lf_mw_Bio$Value=as.numeric(CE_lf_mw_Bio2050$Value)*CR$CR[90]/(25*(1-CR$CR[90])*(1-0.05))*0.048*0.35*Elc_CI$CI/1000#CCP_lf单位为kg, 天然气热值为0.048TJ/ton，发电效率0.35，ELC_CI单位为kg CO2/TJ，最后结果单位为kg
  AE_lf_mw_Bio[is.na(AE_lf_mw_Bio)]=0
  
  ####################################### 5.3 生活废物填埋逸出的CO2（CESC_lf_mw_Bio）#####################################################
  ########排除了非生物质##########
  CESC_lf_mw_Bio=CE_lf_mw_Bio2050[,1:2]
  CESC_lf_mw_Bio$Value=(as.numeric(CE_lf_mw_Bio2050$Value)/25)*(0.05/(1-0.05))*44/16#最终单位为kg
  
  ####################################### 5.4碳汇——（选用CS_lf_mw_Bio）######
  ########使用每年的填埋处理总量（inflow）减掉释放出去的部分（Outflow）得到当年的填埋存量净增量############
  ##假设1961年之前的存量为0
  ########排除了非生物质###########
  CS_lf_mw_Bio=CESC_lf_mw_Bio
  for (n in 1:30){
    CS_lf_mw_Bio[n,1]=RG[n,2]
    CS_lf_mw_Bio$Value[n]=(subset(Landfill_6150,Year==2050&Region==RG[n,2])$Value*0.45*1000-as.numeric(CO2_lf_mw_Bio2050$Value[n])*12/44-
                             as.numeric(CE_lf_mw_Bio2050$Value[n])*12/(25*16)-CECP_lf_mw_Bio$Value[n]*12/44-
                             CESC_lf_mw_Bio$Value[n]*12/44)*44/12#最终单位为kg
  }
  
  
  ####################################### 6. 生活废物非能源回收所含碳################
  ########排除了生物质###########
  CS_ner_mw_Bio=CESC_lf_mw_Bio
  for (n in 1:30){
    CS_ner_mw_Bio$Value[n]=SF30_Bio$Value[n]*0.45*1000*44/12#最终单位为kg CO2
  }
  
  
  
  #######################################（三）工厂废物去向 ##########################################################
  ####################################### 1.工业废物能源回收 #########################
  ####################################### 1.1工业废物能源回收部分的碳排放（CE_er_iw）################################################################
  ###########本来就不包含非生物质############
  CE_er_iw=CS_ner_mw_Bio
  CE_er_iw$Value=SF32$Value*1000*0.45*44/12
  
  
  ####################################### 1.2 工业废物回收的能源（选用E_er_iw）################################################################
  ###########本来就不包含非生物质############
  E_er_iw=CE_er_iw
  E_er_iw$Value=SF32$Value*1000*0.45*0.0156/0.45
  
  
  ####################################### 1.3 工业废物能源回收避免的碳排放（选用AE_er_iw）#########################
  ###########本来就不包含非生物质############
  AE_er_iw=E_er_iw
  AE_er_iw$Value=SF32$Value*1000*0.45*0.0156/0.45*0.25*Elc_CI$CI/1000#最终单位为kg CO2
  
  
  
  #######################################（五）生产过程能耗与碳排放 ############################################
  ####################################### 1. 生产过程直接能耗(自下而上：单耗*产量) #########################################
  ProductPRD_prt=ProductPRD
  ProductPRD_prt$PR=ProductPRD$NP+ProductPRD$PW+ProductPRD$PP*0.17#假设出来需要打印的产品重量
  ProductPRD_prt_L=melt(ProductPRD_prt,id=c("Region"),variable.name = "Product",value.name = "Value")#再次转换为长数据，便于计算
  #######################################【参数修改点5】生产单耗SEC: e #####################################################
  SEC2050_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "SEC_2050")#读入2019年的单耗数据
  SEC2050_Data_L=melt(SEC2050_Data,id=c("Region","Scenario"),variable.name = "Product",value.name = "Value")#将原始宽型数据重组为长型数据
  SEC2050_Data_L=SEC2050_Data_L[order(SEC2050_Data_L$Product),]
  SEC2050_ScName=unique(SEC2050_Data_L$Scenario)
  SEC2050=subset(SEC2050_Data_L,Scenario==SEC2050_ScName[e])
  E_prd=SEC2050
  E_prd$Value=ProductPRD_prt_L$Value*SEC2050$Value/1000000000#单位为GJ
  
  
  ####################################### 2. 生产过程总的直接能耗 #########################################################
  E_prd_W=dcast(E_prd,Region~Product,value.var = "Value")
  E_prd_W$SUM=E_prd_W$CWP+E_prd_W$HS+E_prd_W$MP+
    E_prd_W$NP+E_prd_W$NWP+E_prd_W$OP+E_prd_W$PP+
    E_prd_W$PW+E_prd_W$RP+E_prd_W$PR
  E_prd_P=E_prd_W[,-(2:11)]
  names(E_prd_P)=c("Region","Value")
  
  ####################################### 3. 估算生物质能源 #####################################################################
  ########################通过备料损失率和工业废物能源回收推算得到的生物质能源(初级能源)####################
  ##计算这一步之前，必须先把工业废物供能计算出来
  Wood=SF1
  Wood$Value=SF1$Value+SF2$Value
  Otherfibre=SF3
  Otherfibre$Value=SF3$Value
  E_er_iw_EW=subset(E_er_iw,Region!=0)
  E_prd_BioWst_estm=E_er_iw_EW
  E_prd_BioWst_estm$Value=Wood$Value*0.056*15.6/(1-0.056)+Otherfibre$Value*0.056*14.7/(1-0.056)+as.numeric(E_er_iw_EW$Value)#estm尾缀代表估算值、推算值;0.056是损失率，15.6GJ/t
  ##计算木材边角料供能
  E_prd_BioWst_estm$WoodE=Wood$Value*0.056*15.6/(1-0.056)+Otherfibre$Value*0.056*14.7/(1-0.056)#14.7GJ/Ton来自中国能源统计年鉴各种能源折标煤参考系数，是四种秸秆的均值
  E_prd_BioWst_estm$iwer=as.numeric(E_er_iw_EW$Value)
  
  ####################说明：未来模拟不需要具体知道哪些产品使用了生物质能和非生物能各多少，因此不进行分配#######################
  
  ####################################### 4. 分能源结构的直接能耗 ######################################################################
  #########【重要参数】能源结构的设置####################
  ##参考情景中，读入2019年的能源结构
  E_prd_Bio=E_prd_P
  E_prd_Bio$Value=E_prd_BioWst_estm$Value*0.85
  E_prd_ExBio=E_prd_Bio
  E_prd_ExBio$Value=E_prd_P$Value-E_prd_Bio$Value
  E_prd_ExBio$Value[E_prd_ExBio$Value<0]=0
  ###################################### 【参数修改点6】能源结构 ES: f ################################################################
  ES_ExBio_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "ES2050")#ES_ExBio是基于IEA数据计算的、除去生物质能源的直接能耗的能源结构比例；埃及和马来西亚单独找了工业层面的能源消费结构数据补充上的
  ES_ExBio_ScName=unique(ES_ExBio_Data$Scenario)
  ES_ExBio=subset(ES_ExBio_Data,Scenario==ES_ExBio_ScName[f])##读取情景数据
  E_prd_ExBio_ES=ES_ExBio
  E_prd_ExBio_ES[,3:12]=E_prd_ExBio$Value*ES_ExBio[,3:12]
  E_prd_ES=E_prd_ExBio_ES
  E_prd_ES$Wood=E_prd_BioWst_estm$WoodE
  E_prd_ES$Black.liquid=E_prd_BioWst_estm$iwer
  
  ####################################### 5. 分能源结构的初级能耗 ######################################################################
  ETF=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "ETF2050")#ETF是分能源类型的转换效率
  ETF=subset(ETF,Region!="World")
  E_prd_ES_Pri=E_prd_ES[,3:12]/ETF[,2:11]
  E_prd_ES_Pri=cbind(E_prd_W[,1],E_prd_ES_Pri)#加上年份和国家
  colnames(E_prd_ES_Pri)[1]="Region"
  
  ####################################### 6. 总的初级能耗 ######################################################################
  E_prd_Pri=E_prd_ES_Pri[,1:2]#本来只要一列，但是这样不能形成数据框，所以先导入两列，后面再删掉第二列
  E_prd_Pri$Value=rowSums(E_prd_ES_Pri[,2:11])
  E_prd_Pri=E_prd_Pri[,-2]#去掉第二列
  
  
  ####################################### 7. 生产过程碳排放计算 ###############################################################
  ##在SC_CarbonEmission_Best-R0-HighCR表格中新建一个sheet，名字为E_prd_BioOrNot_AllEstm，将生产过程的总初级能耗按照生物质和非生物质分开
  E_prd_BioOrNot_AllEstm=E_prd_Pri
  E_prd_BioOrNot_AllEstm$E_prd_Bio=E_prd_BioWst_estm$Value
  E_prd_BioOrNot_AllEstm$E_prd_ExBio=E_prd_Pri$Value-E_prd_BioWst_estm$Value
  E_prd_BioOrNot_AllEstm=E_prd_BioOrNot_AllEstm[,-2]
  
  
  ############### 8. 生产过程——碳排(自下而上)#############################
  #####################################【参数修改点7】能源排放系数：g ################################################
  ##导入各能源二氧化碳含量
  CC2050_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName = "CO2_Con2050")
  CC2050_ScName=unique(CC2050_Data$Scenario)
  CC2050=subset(CC2050_Data,Scenario==CC2050_ScName[g])
  ##计算生产过程能源相关碳排放
  CE_prd=E_prd_ES_Pri[,2:11]*CC2050[,3:12]/1000#单位是kg CO2
  CE_prd=cbind(E_prd_ES_Pri$Region,CE_prd)
  names(CE_prd)=c(colnames(E_prd_ES_Pri))##这一句不加也可以，原本列名就是正确的
  CE_prd$SUM=rowSums(CE_prd[,2:11])#各列求和
  CE_prd$Bio=rowSums(CE_prd[,10:11])#生物质能源碳排放（包括工业废物、生活废物、生物质燃料）
  CE_prd$ExBio=CE_prd$SUM-CE_prd$Bio#化石能源碳排放
  
  
  ##以下将总量、生物质和非生物质碳排分别存入表格备用
  CE_prd_SUM=CE_prd[,1:2]
  CE_prd_SUM$Value=CE_prd$SUM
  CE_prd_SUM=CE_prd_SUM[,-2]
  
  
  CE_prd_Bio=CE_prd[,1:2]
  CE_prd_Bio$Value=CE_prd$Bio
  CE_prd_Bio=CE_prd_Bio[,-2]
  
  
  CE_prd_ExBio=CE_prd[,1:2]
  CE_prd_ExBio$Value=CE_prd$ExBio
  CE_prd_ExBio=CE_prd_ExBio[,-2]
  AE_prd=CE_prd_ExBio
  AE_prd$Value[AE_prd$Value>0]=0
  CE_prd_ExBio[CE_prd_ExBio<0]=0
  #######################################（三）工厂废物去向2 ##########################################################
  ####################################### 0. 比较特殊：边角料能源回收 ###################################################
  ####################################### 0.1 边角料产生的碳排 ###################################################
  CE_WoodE=E_prd_BioWst_estm
  CE_WoodE=CE_WoodE[,-(3:4)]
  CE_WoodE$Value=CE_prd$Wood
  
  ####################################### 0.2 边角料回收的能源 ################################################################
  E_WoodE=CE_WoodE
  E_WoodE$Value=CE_WoodE$Value*(12*0.0147)/(0.45*44)##边角料低位热值为0.0147
  
  ####################################### 0.3 边角料能源回收避免的碳排放 ################################################################
  ###########本来就不包含非生物质############
  AE_WoodE=E_WoodE
  AE_WoodE$Value=E_WoodE$Value*0.25*Elc_CI$CI/1000##0.25是转化率
  
  #######################################（四）不可持续采伐引起的毁林碳排放 ###################################################################################
  ####################################### 【参数修改点8】森林碳排放情景模块：h ######################################
  WoodPerPw2019=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="WoodPerPw2019")
  ForestCarbonEF2050_Data=read.xlsx("INPUT/Scenario/ScenarioModelling_Info.xlsx",sheetName="ForestCarbonEF2050")
  ForestCarbonEF2050_ScName=unique(ForestCarbonEF2050_Data$Scenario)
  ForestCarbonEF2050=subset(ForestCarbonEF2050_Data,Scenario==ForestCarbonEF2050_ScName[h])
  RWPP=WoodPRD
  RWPP$Value= ifelse(ForestCarbonEF2050$WPP_FPE/(WoodPRD$Value*WoodPerPw2019$R_RWtoPw*ForestCarbonEF2050$R_IRWfrNF_USCC*ForestCarbonEF2050$R_Usus)>1,1,
                    ForestCarbonEF2050$WPP_FPE/(WoodPRD$Value*WoodPerPw2019$R_RWtoPw*ForestCarbonEF2050$R_IRWfrNF_USCC*ForestCarbonEF2050$R_Usus))
  RWPP[is.na(RWPP)]=0
  RCC=RWPP
  RCC$Value=ifelse(RWPP$Value>ForestCarbonEF2050$R_IRWfrNFbyCC_USCC,0,ForestCarbonEF2050$R_IRWfrNFbyCC_USCC-RWPP$Value)
  RCC[is.na(RCC)]=0
  RSL=RCC
  RSL$Value=ifelse(RWPP$Value==1,0,ifelse(RWPP$Value>ForestCarbonEF2050$R_IRWfrNFbyCC_USCC,1-RWPP$Value,1-ForestCarbonEF2050$R_IRWfrNFbyCC_USCC))
  RSL[is.na(RSL)]=0
  
  CE_df=WoodPRD##WoodPRD$Value
  CE_df$Value=(WoodPRD$Value*(ForestCarbonEF2050$TBC_FPE/(WoodPRD$Value*WoodPerPw2019$R_RWtoPw)+ForestCarbonEF2050$PC_FPE)+##forest plantation expansion
    WoodPRD$Value*(ForestCarbonEF2050$R_IRWfrNF_USCC* RCC$Value*ForestCarbonEF2050$CF/ForestCarbonEF2050$CutPerHa)*
      (ForestCarbonEF2050$C.stocks.of.forest+ForestCarbonEF2050$SCL_USCC)*(44/12)*ForestCarbonEF2050$R_Usus+##CC碳排放
    WoodPRD$Value*RSL$Value*ForestCarbonEF2050$R_IRWfrNF_USCC*ForestCarbonEF2050$ESofSL_USSL*(44/12)*ForestCarbonEF2050$R_Usus-##SL碳排放
    WoodPRD$Value*(ForestCarbonEF2050$R_IRWfrNF_USCC*((RCC$Value+RSL$Value)*ForestCarbonEF2050$R_Usus))*0.5*0.45*(44/12))*##减掉CC和SL多算的木材中的碳
    1000##森林碳这一块，在excel中计算结果的单位都是tCO2，而R中计算都是kg，因此这里要乘以1000换算成kg
  CE_df[is.na(CE_df)]=0
  
  ColumnName=paste(pq,"_",Elc_CI_ScName[a],"_",DMD_ScName[b],"_",
                   RcyclRate2050_ScName[c],"_",YR_ScName[d],"_",SEC2050_ScName[e],"_",
                   ForestCarbonEF2050_ScName[h],"_",CR_Name[l],"(",pq,"_",a,b,c,d,e,h,l,")",sep="",collapse = NULL)
  
  filename=paste("OUT/Scenario/All scenarios (2160)/",Elc_CI_ScName[a],"_",DMD_ScName[b],"_",
                 RcyclRate2050_ScName[c],"_",YR_ScName[d],"_",SEC2050_ScName[e],"_",
                 ForestCarbonEF2050_ScName[h],"_",CR_Name[l],"(",pq,"_",a,b,c,d,e,h,l,")",".xlsx",sep="",collapse = NULL)
  write.xlsx(CE_wttrt,filename,sheetName="CE_wttrt",append=TRUE)#导出数据15 
  write.xlsx(CE_wttrt_ToNature,filename,sheetName="CE_wttrt_ToNature",append=TRUE)
  write.xlsx(CE_wttrt_Pulp_SUM,filename,sheetName="CE_wttrt_Pulp",append=TRUE)
  write.xlsx(CE_wttrt_Paper_SUM,filename,sheetName="CE_wttrt_Paper",append=TRUE)#导出数据15
  write.xlsx(CE_wttrt_ToNature_Pulp_SUM,filename,sheetName="CE_wttrt_ToNature_Pulp",append=TRUE)#导出数据15
  write.xlsx(CE_wttrt_ToNature_Paper_SUM,filename,sheetName="CE_wttrt_ToNature_Paper",append=TRUE)#导出数据15
  write.xlsx(E_wh,filename,sheetName="E_wh",append=TRUE)
  write.xlsx(CE_wh,filename,sheetName="CE_wh",append=TRUE)
  write.xlsx(E_nwh,filename,sheetName="E_nwh",append=TRUE)#导出数据5
  write.xlsx(CE_nwh,filename,sheetName="CE_nwh",append=TRUE)#导出数据6
  write.xlsx(E_rpc,filename,sheetName="E_rpc",append=TRUE)
  write.xlsx(CE_rpc,filename,sheetName="CE_rpc",append=TRUE)
  write.xlsx(E_he,filename,sheetName="E_he",append=TRUE)
  write.xlsx(CE_he,filename,sheetName="CE_he",append=TRUE)
  write.xlsx(E_chem,filename,sheetName="E_chem",append=TRUE)
  write.xlsx(CE_chem,filename,sheetName="CE_chem",append=TRUE)
  write.xlsx(CS_ps_Bio,filename,sheetName="CS_ps_Bio",append=TRUE)
  write.xlsx(CE_inc_Bio,filename,sheetName="CE_inc_Bio",append=TRUE)#导出数据20
  write.xlsx(CE_er_mw_Bio,filename,sheetName="CE_er_mw_Bio",append=TRUE)#导出数据21
  write.xlsx(E_er_mw,filename,sheetName="E_er_mw",append=TRUE)
  write.xlsx(AE_er_mw,filename,sheetName="AE_er_mw",append=TRUE)#导出
  write.xlsx(CO2_lf_mw_Bio,filename,sheetName="CO2_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(CO2_lf_mw_Bio2050,filename,sheetName="CO2_lf_mw_Bio2050",append=TRUE)#导出数据
  write.xlsx(CE_lf_mw_Bio,filename,sheetName="CE_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(CE_lf_mw_Bio2050,filename,sheetName="CE_lf_mw_Bio2050",append=TRUE)#导出数据
  write.xlsx(CECP_lf_mw_Bio,filename,sheetName="CECP_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(E_lf_mw_Bio,filename,sheetName="E_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(AE_lf_mw_Bio,filename,sheetName="AE_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(CESC_lf_mw_Bio,filename,sheetName="CESC_lf_mw_Bio",append=TRUE)#导出数据28
  write.xlsx(CS_lf_mw_Bio,filename,sheetName="CS_lf_mw_Bio",append=TRUE)#导出数据
  write.xlsx(CS_ner_mw_Bio,filename,sheetName="CS_ner_mw_Bio",append=TRUE)#导出数据
  write.xlsx(CE_er_iw,filename,sheetName="CE_er_iw",append=TRUE)
  write.xlsx(E_er_iw,filename,sheetName="E_er_iw",append=TRUE)
  write.xlsx(AE_er_iw,filename,sheetName="AE_er_iw",append=TRUE)
  write.xlsx(E_prd,filename,sheetName="E_prd",append=TRUE)#E_prd是生产阶段带有产品结构的直接能耗
  write.xlsx(E_prd_P,filename,sheetName="E_prd_P",append=TRUE)#E_prd_P是生产阶段总的直接能耗
  write.xlsx(E_prd_BioWst_estm,filename,sheetName="E_prd_BioWst_estm",append=TRUE)#E_prd_BioWst_estm是估算的生物质能源（初级）供应量
  write.xlsx(E_prd_ES,filename,sheetName="E_prd_ES",append=TRUE)#E_prd_ES是分能源类型结构的直接能耗
  write.xlsx(E_prd_ES_Pri,filename,sheetName="E_prd_ES_Pri",append=TRUE)#E_prd_ES是分能源类型结构的直接能耗
  write.xlsx(E_prd_Pri,filename,sheetName="E_prd_Pri",append=TRUE)#E_prd_Pri是IEA能源结构推算得到的总初级能耗
  write.xlsx(E_prd_BioOrNot_AllEstm,filename,sheetName="E_prd_BioOrNot_AllEstm",append=TRUE)#E_prd_BioWst_estm是估算的生物质能源（初级）供应量
  write.xlsx(CE_prd,filename,sheetName="CE_prd",append=TRUE)#分能源结构的碳排放（基于初级能耗）
  write.xlsx(CE_prd_SUM,filename,sheetName="CE_prd_SUM",append=TRUE)#排放总值
  write.xlsx(CE_prd_Bio,filename,sheetName="CE_prd_Bio",append=TRUE)#生物质能源排放
  write.xlsx(CE_prd_ExBio,filename,sheetName="CE_prd_ExBio",append=TRUE)#非生物质能源排放
  write.xlsx(AE_prd,filename,sheetName="AE_prd",append=TRUE)##制浆造纸工业用不了的生物质能源导致的负排放（利用方式同制浆造纸行业，能源结构以情景设置为准）
  write.xlsx(CE_WoodE,filename,sheetName="CE_WoodE",append=TRUE)
  write.xlsx(E_WoodE,filename,sheetName="E_WoodE",append=TRUE)
  write.xlsx(AE_WoodE,filename,sheetName="AE_WoodE",append=TRUE)#导出数据24
  write.xlsx(CE_df,filename,sheetName="CE_df",append=TRUE)#导出数据24
  
  ####################################### 三、针对分产品种类的项进行求和 #######################################
  TN1=c( "CE_chem","CE_wttrt","E_chem")
  for (i in 1:3){
    data0=read.xlsx(filename,sheetName = TN1[i])
    data0$Value=as.numeric(data0$Value)
    data_P=dcast(data0,Region~Product,value.var = "Value")
    NC=ncol(data_P)
    data_PA=data_P[,2:NC]
    data_T=data_P[,1:2]
    data_T$Value=apply(data_PA,1,sum)
    data_T=data_T[,-2]
    write.xlsx(data_T,filename,sheetName=paste(TN1[i],"_P",sep = "",collapse = NULL),append=TRUE)
  }
  
  ####################################### 四、碳排放和碳汇构造表格 ###################################################
  ############################## 1. 生命周期碳排放和碳汇构造表格 ###################################################
  TN2=c("CS_ps_Bio", "CS_lf_mw_Bio","CS_ner_mw_Bio",
        "AE_er_mw","AE_lf_mw_Bio","AE_er_iw","AE_WoodE","AE_prd",
        "CE_df","CE_wh","CE_nwh","CE_rpc","CE_chem_P","CE_wttrt_Pulp","CE_wttrt_Paper","CE_wttrt_ToNature_Pulp","CE_wttrt_ToNature_Paper",
        "CE_prd_ExBio","CE_prd_Bio",
        "CE_inc_Bio","CE_er_mw_Bio", "CO2_lf_mw_Bio2050","CE_lf_mw_Bio2050","CECP_lf_mw_Bio","CESC_lf_mw_Bio",
        "CE_er_iw","CE_WoodE")
  CE_chem_P=read.xlsx(filename,"CE_chem_P")
  Carbon=SF21
  Carbon$CS_ps_Bio=-CS_ps_Bio$Value
  Carbon$CS_lf_mw_Bio=-CS_lf_mw_Bio$Value
  Carbon$CS_ner_mw_Bio=-CS_ner_mw_Bio$Value
  Carbon$AE_er_mw=-AE_er_mw$Value
  Carbon$AE_lf_mw_Bio=-AE_lf_mw_Bio$Value
  Carbon$AE_WoodE=-AE_WoodE$Value
  Carbon$AE_prd=AE_prd$Value
  Carbon$CE_df=CE_df$Value
  Carbon$CE_wh=CE_wh$Value
  Carbon$CE_nwh=CE_nwh$Value
  Carbon$CE_rpc=CE_rpc$Value
  Carbon$CE_chem_P=CE_chem_P$Value
  Carbon$CE_wttrt_Pulp=CE_wttrt_Pulp_SUM$Value
  Carbon$CE_wttrt_Paper=CE_wttrt_Paper_SUM$Value
  Carbon$CE_wttrt_ToNature_Pulp=CE_wttrt_ToNature_Pulp_SUM$Value
  Carbon$CE_wttrt_ToNature_Paper=CE_wttrt_ToNature_Paper_SUM$Value
  Carbon$CE_prd_ExBio=CE_prd_ExBio$Value
  Carbon$CE_prd_Bio=CE_prd_Bio$Value
  Carbon$CE_inc_Bio=CE_inc_Bio$Value
  Carbon$CE_er_mw_Bio=CE_er_mw_Bio$Value
  Carbon$CO2_lf_mw_Bio2050=CO2_lf_mw_Bio2050$Value
  Carbon$CE_lf_mw_Bio2050=CE_lf_mw_Bio2050$Value
  Carbon$CECP_lf_mw_Bio=CECP_lf_mw_Bio$Value
  Carbon$CESC_lf_mw_Bio=CESC_lf_mw_Bio$Value
  Carbon$CE_er_iw=CE_er_iw$Value
  Carbon$CE_WoodE=CE_WoodE$Value
  Carbon_num=as.data.frame(lapply( Carbon,as.numeric))
  Carbon[2:16]=Carbon_num[2:16]
  Carbon=Carbon[,-(2)]
  write.xlsx(Carbon,filename,sheetName="Carbon",append = TRUE)
  
  ############################## 2. 重分类、绘图前准备等处理 ###################################################
  ############################## 2.1 所有来源的碳排放都计入 ###################################################
  ############################## 2.1.1 重分类 ###################################################################
  Carbon_ReCls= Carbon[,1:2]##定义第一列地区，由于少于2列无法形成数据框，所以只能选2列，后面再去掉
  ##定义其后每一列数据
  Carbon_ReCls$CS_ps=Carbon$CS_ps_Bio
  Carbon_ReCls=Carbon_ReCls[,-2]#去掉被迫留着的第二列
  Carbon_ReCls$CS_lf=Carbon$CS_lf_mw_Bio
  Carbon_ReCls$CS_ner=Carbon$CS_ner_mw_Bio-Carbon$CE_wttrt_ToNature_Pulp-Carbon$CE_wttrt_ToNature_Paper
  Carbon_ReCls$AE_er=Carbon$AE_er_mw
  Carbon_ReCls$AE_lf=Carbon$AE_lf_mw_Bio-(CE_wttrt_Paper_SUM$Value+CE_wttrt_Pulp_SUM$Value)*CR$CR[90]*0.048*0.35*Elc_CI$CI/(25*1000)
  Carbon_ReCls$AE_prd=Carbon$AE_prd
  Carbon_ReCls$CE_df=Carbon$CE_df
  Carbon_ReCls$CE_he=Carbon$CE_wh+Carbon$CE_nwh+Carbon$CE_rpc
  Carbon_ReCls$CE_chem=Carbon$CE_chem_P
  Carbon_ReCls$CE_wttrt=(Carbon$CE_wttrt_Pulp+Carbon$CE_wttrt_Paper)*(1-CR$CR[90])
  Carbon_ReCls$CE_prd_EXBio=Carbon$CE_prd_ExBio
  Carbon_ReCls$CE_inc=Carbon$CE_inc_Bio
  Carbon_ReCls$CE_er=Carbon$CE_er_mw_Bio+Carbon$CE_er_iw+Carbon$CE_WoodE
  Carbon_ReCls$CE_lf=Carbon$CE_lf_mw_Bio
  Carbon_ReCls$CECP_lf=Carbon$CECP_lf_mw_Bio+
    (CE_wttrt_Pulp_SUM$Value+CE_wttrt_Paper_SUM$Value)*CR$CR[90]*44/(25*16)
  Carbon_ReCls$CESC_lf=as.numeric(Carbon$CO2_lf_mw_Bio)+as.numeric(Carbon$CESC_lf_mw_Bio)
  ##增加三列，分别是碳汇、碳排和净碳排
  Carbon_ReCls_num=as.data.frame(lapply( Carbon_ReCls,as.numeric))
  Carbon_ReCls[2:17]=Carbon_ReCls_num[2:17]
  Carbon_ReCls$CS=rowSums(Carbon_ReCls[,2:7])/10^9##CE表示Carbon Sink or Carbon Stock
  Carbon_ReCls$CE=rowSums(Carbon_ReCls[,8:17])/10^9##CE表示是Carbon Emission
  Carbon_ReCls$NCE=Carbon_ReCls$CE+Carbon_ReCls$CS##NCE表示Net Carbon Emission
  Carbon_ReCls$CS_ExAE=Carbon_ReCls$CS-(Carbon_ReCls$AE_er+Carbon_ReCls$AE_lf+Carbon_ReCls$AE_prd)/10^9
  Carbon_ReCls$NCE_ExAE=Carbon_ReCls$CE+Carbon_ReCls$CS_ExAE##NCE表示Net Carbon Emission
  Carbon_ReCls$CS_ExAEprd=Carbon_ReCls$CS-(Carbon_ReCls$AE_prd)/10^9
  Carbon_ReCls$NCE_ExAEprd=Carbon_ReCls$CE+Carbon_ReCls$CS_ExAEprd##NCE表示Net Carbon Emission
  Carbon_ReCls[is.na(Carbon_ReCls)]=0
  ##导出数据
  write.xlsx(Carbon_ReCls,filename,sheetName="Carbon_ReCls",append = TRUE)##导出数据
  
  ############################## 2.2 生物质来源的碳排放视为碳中性 ###################################################
  ##建立SC_Carbon_Best-R0-HighCR工作簿的副本，改名为SC_Carbon_BCN
  ############################## 2.2.1 重分类 ###################################################################
  Carbon_ReCls_BCN= Carbon[,1:2]##定义前两列地区和年份
  ##定义其后每一列数据
  Carbon_ReCls_BCN$CS_ps=Carbon$CS_ps_Bio
  Carbon_ReCls_BCN=Carbon_ReCls_BCN[,-2]#去掉被迫留着的第二列
  Carbon_ReCls_BCN$CS_lf=Carbon$CS_lf_mw_Bio
  Carbon_ReCls_BCN$CS_ner=Carbon$CS_ner_mw_Bio-Carbon$CE_wttrt_ToNature_Pulp-Carbon$CE_wttrt_ToNature_Paper
  Carbon_ReCls_BCN$AE_er=Carbon$AE_er_mw
  Carbon_ReCls_BCN$AE_lf=Carbon$AE_lf_mw_Bio-(CE_wttrt_Paper_SUM$Value+CE_wttrt_Pulp_SUM$Value)*CR$CR[90]*0.048*0.35*Elc_CI$CI/(25*1000)
  Carbon_ReCls_BCN$AE_prd=Carbon$AE_prd
  Carbon_ReCls_BCN$CE_df=Carbon$CE_df
  Carbon_ReCls_BCN$CE_he=Carbon$CE_wh+Carbon$CE_nwh+Carbon$CE_rpc
  Carbon_ReCls_BCN$CE_chem=Carbon$CE_chem_P
  Carbon_ReCls_BCN$CE_wttrt=(Carbon$CE_wttrt_Pulp+Carbon$CE_wttrt_Paper)*(1-CR$CR[90])
  Carbon_ReCls_BCN$CE_prd_EXBio=Carbon$CE_prd_ExBio
  Carbon_ReCls_BCN$CE_inc=Carbon$CE_inc_Bio
  Carbon_ReCls_BCN$CE_er=Carbon$CE_er_mw_Bio+Carbon$CE_er_iw+Carbon$CE_WoodE
  Carbon_ReCls_BCN$CE_lf=Carbon$CE_lf_mw_Bio
  Carbon_ReCls_BCN$CECP_lf=Carbon$CECP_lf_mw_Bio+(CE_wttrt_Pulp_SUM$Value+CE_wttrt_Paper_SUM$Value)*CR$CR[90]*44/(25*16)
  Carbon_ReCls_BCN$CESC_lf=as.numeric(Carbon$CO2_lf_mw_Bio)+as.numeric(Carbon$CESC_lf_mw_Bio)
  ##增加三列，分别是碳汇、碳排和净碳排
  Carbon_ReCls_num=as.data.frame(lapply( Carbon_ReCls,as.numeric))
  Carbon_ReCls[2:17]=Carbon_ReCls_num[2:17]
  Carbon_ReCls_BCN$CS=rowSums(Carbon_ReCls_BCN[,2:7])/10^9##CS表示Carbon Sink or Carbon Stock
  Carbon_ReCls_BCN$CE=(Carbon_ReCls_BCN$CE_df+Carbon_ReCls_BCN$CE_he+Carbon_ReCls_BCN$CE_chem+Carbon_ReCls_BCN$CE_wttrt+
                         Carbon_ReCls_BCN$CE_prd_EXBio+as.numeric(Carbon_ReCls_BCN$CE_lf))/10^9##CE表示是Carbon Emission
  Carbon_ReCls_BCN$NCE=Carbon_ReCls_BCN$CE+Carbon_ReCls_BCN$CS##NCE表示Net Carbon Emission
  Carbon_ReCls_BCN$CS_ExAE=Carbon_ReCls_BCN$CS-(Carbon_ReCls_BCN$AE_er+Carbon_ReCls_BCN$AE_lf+Carbon_ReCls_BCN$AE_prd)/10^9
  Carbon_ReCls_BCN$NCE_ExAE=Carbon_ReCls_BCN$CE+Carbon_ReCls_BCN$CS_ExAE##NCE表示Net Carbon Emission
  Carbon_ReCls_BCN$CS_ExAEprd=Carbon_ReCls_BCN$CS-(Carbon_ReCls_BCN$AE_prd)/10^9
  Carbon_ReCls_BCN$NCE_ExAEprd=Carbon_ReCls_BCN$CE+Carbon_ReCls_BCN$CS_ExAEprd##NCE表示Net Carbon Emission
  Carbon_ReCls_BCN[is.na(Carbon_ReCls_BCN)]=0##缺失值替换为0
  ##导出数据
  write.xlsx(Carbon_ReCls_BCN,filename,sheetName="Carbon_ReCls_BCN",append = TRUE)##导出数据
  ##导出数据
  SC_n1=Carbon_ReCls_BCN$NCE
  SC_n1=data.frame(SC_n1)
  names(SC_n1)=ColumnName
  AllScs=cbind(AllScs,SC_n1)
  ##导出数据
  SC_n2=Carbon_ReCls_BCN$NCE_ExAE
  SC_n2=data.frame(SC_n2)
  names(SC_n2)=ColumnName
  AllScs_ExAE=cbind(AllScs_ExAE,SC_n2)
  ##导出数据
  SC_n3=Carbon_ReCls_BCN$NCE_ExAEprd
  SC_n3=data.frame(SC_n3)
  names(SC_n3)=ColumnName
  AllScs_ExAEprd=cbind(AllScs_ExAEprd,SC_n3)
}

##单独导出净排放数据
#write.xlsx(AllScs,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE",append = TRUE)
write.xlsx(AllScs_ExAE,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAE",append = TRUE)
#write.xlsx(AllScs_ExAEprd,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAEprd",append = TRUE)
##读入并导出长数据
#AllScs=read.xlsx("OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE")
#AllScs=AllScs[,-1]
#AllScs_L=melt(AllScs,id=c("Region","Abbreviation"),variable.name = "Scenario",value.name = "Value")
#write.xlsx(AllScs_L,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_L",append = TRUE)

AllScs_ExAE=read.xlsx("OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAE")
AllScs_ExAE=AllScs_ExAE[,-1]
AllScs_ExAE_L=melt(AllScs_ExAE,id=c("Region","Abbreviation"),variable.name = "Scenario",value.name = "Value")
write.xlsx(AllScs_ExAE_L,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAE_L",append = TRUE)

#AllScs_ExAEprd=read.xlsx("OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAEprd")
#AllScs_ExAEprd=AllScs_ExAEprd[,-1]
#AllScs_ExAEprd_L=melt(AllScs_ExAEprd,id=c("Region","Abbreviation"),variable.name = "Scenario",value.name = "Value")
#write.xlsx(AllScs_ExAEprd_L,"OUT/Scenario/All scenarios (2160)/AllScs.xlsx",sheetName = "AllScs_NCE_ExAEprd_L",append = TRUE)
