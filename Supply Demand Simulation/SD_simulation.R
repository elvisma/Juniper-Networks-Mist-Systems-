####### Juniper/Mist SD_running simulation ########
#install.packages("xlsx")

library("xlsx")
library("formattable")
library("openxlsx")
library("readxl")
library("tidyr")
options(scipen=999)
library("dplyr")


setwd("C:/Users/xiaoxiongm/Desktop/Projects/Actuals v.s. Forecasts")


#### get demand ####

##### 1. get backlog #############


#"New Order Item Status", "Product Code", "QTY", "CRDD YEAR", "WK reference","no-ship QTY", "Line Category General", "Line Category", "Warehouse"
order_df<-read_excel("historical v.s. Forecast waterfall (draft).xlsx",sheet="order_site_CRDD")[,c(25,19,20,29,35,37:40)]

'%!in%' <- function(x,y)!('%in%'(x,y))
backlog_df=order_df[which(order_df$`Line Category`%!in%c("Cancelled","Shipped")),]
backlog_df=backlog_df[which(backlog_df$`New Order Item Status`!="Draft"),]
backlog_df$`WK reference`<-as.Date(backlog_df$`WK reference`)
backlog_df[is.na(backlog_df)]=0
backlog_summ=spread(summarise(group_by(backlog_df,Warehouse,`Product Code`,`WK reference`,`Line Category`),QTY=sum(`no-ship QTY`)),'Line Category','QTY',fill = 0)


###### 2. get commit, upside, pipeline #########
total_deal_col_types=c("text","text","text","text","text","text","numeric","text","text","text","text","text","text","numeric","numeric","numeric","numeric","numeric","numeric","text","text","text","text","text","text")
commit_df<-read_excel("historical v.s. Forecast waterfall (draft).xlsx",sheet="total_deals (with Pipeline)",col_types=total_deal_col_types,skip=7)[,c(1,3,4,7,13,17,18,19,25)]
opp_df <-commit_df[which(commit_df$`Line Category`%in%c("Commit","Upside","Pipeline")),]
opp_df$`WK reference`<-as.Date(opp_df$`WK reference`)
opp_df[is.na(opp_df)]=0
opp_summ<-summarise(group_by(opp_df,Warehouse, `Product Code`, `WK reference`),Commit=sum(Commit),Upside=sum(Upside),Pipeline=sum(Pipeline))
opp_summ[is.na(opp_summ)]=0

############ combine demand #####################
demand_total=merge(backlog_summ,opp_summ,all=TRUE)
demand_total[is.na(demand_total)]=0
colnames(demand_total)[1:3]=c("WH","SKU","WK")
#demand_total= demand_total[,c("WH","SKU","WK","cmtd BKLG (past Schedule)","cmtd BKLG (future Schedule)","non-cmtd BKLG (past CRDD)","non-cmtd BKLG (future CRDD)","Commit","Upside","Pipeline")]


#### get supply ####

setwd("C:/Users/xiaoxiongm/Desktop/Projects/Ocean Air Analysis")
supply_df<-read_excel("Ocean Air analysis for Top SKUs.xlsm",sheet="Supplies",skip=3)[,c(2,3,24:36)]
supply_df_active=supply_df[which(supply_df$Status!='Inactive'),]

supply_df_active$`MP Air ETA WK`=as.Date(supply_df_active$`MP Air ETA WK`)
supply_df_active$`MP Ocean ETA WK`=as.Date(supply_df_active$`MP Ocean ETA WK`)
supply_df_active$`AMS Air ETA WK`=as.Date(supply_df_active$`AMS Air ETA WK`)
supply_df_active$`AMS Ocean ETA WK`=as.Date(supply_df_active$`AMS Ocean ETA WK`)
supply_df_active$`HK Air ETA WK`=as.Date(supply_df_active$`HK Air ETA WK`)
supply_df_active$`HK Ocean ETA WK`=as.Date(supply_df_active$`HK Ocean ETA WK`)


########## Milpitas ##############
MP_air_supply_df<-supply_df_active[,c(1,2,3:4)]
MP_air_supply_df<-MP_air_supply_df[which(!is.na(MP_air_supply_df$`MP Air ETA WK`)),]


MP_ocean_supply_df<-supply_df_active[,c(1,2,5:6)]
MP_ocean_supply_df<-MP_ocean_supply_df[which(!is.na(MP_ocean_supply_df$`MP Ocean ETA WK`)),]


#Category (in-transit, ship plan, build plan) pivot
MP_air=MP_air_supply_df
colnames(MP_air)[3:4]=c("QTY","WK")

MP_ocean=MP_ocean_supply_df
colnames(MP_ocean)[3:4]=c("QTY","WK")

MP_category<-rbind(MP_air,MP_ocean)
MP_category$WH<-"MP"
MP_category_group<-group_by(MP_category,WH,SKU,WK,CATEGORY)
MP_category_summ<-summarise(MP_category_group, QTY=sum(QTY))
MP_ctgry<-spread(MP_category_summ,'CATEGORY','QTY',fill=0)

#ship method (Air Ocean) pivot
MP_air_MOT<-MP_air_supply_df
colnames(MP_air_MOT)[4]='WK'
MP_air_MOT <-summarise(group_by(MP_air_MOT,SKU,WK),Air_QTY=sum(`MP Air QTY`))

MP_ocean_MOT<-MP_ocean_supply_df
colnames(MP_ocean_MOT)[4]='WK'
MP_ocean_MOT <-summarise(group_by(MP_ocean_MOT,SKU,WK),Ocean_QTY=sum(`MP Ocean QTY`))

MP_MOT<-merge(MP_air_MOT,MP_ocean_MOT,all = TRUE)
MP_MOT$WH="MP"
MP_MOT[is.na(MP_MOT)]=0

# Merge (category, ship method)

MP_total=merge(MP_ctgry,MP_MOT,all = TRUE)
MP_total=MP_total[,c("WH","SKU","WK","in-transit","ship plan","Air_QTY","Ocean_QTY")]


#############################################################################################################


########## Amsterdam ##############


AMS_air_supply_df<-supply_df_active[,c(1,2,7:8)]
AMS_air_supply_df<-AMS_air_supply_df[which(!is.na(AMS_air_supply_df$`AMS Air ETA WK`)),]

AMS_ocean_supply_df<-supply_df_active[,c(1,2,9:10)]
AMS_ocean_supply_df<-AMS_ocean_supply_df[which(!is.na(AMS_ocean_supply_df$`AMS Ocean ETA WK`)),]



#Category (in-transit, ship plan, build plan) pivot
AMS_air=AMS_air_supply_df
colnames(AMS_air)[3:4]=c("QTY","WK")

AMS_ocean=AMS_ocean_supply_df
colnames(AMS_ocean)[3:4]=c("QTY","WK")

AMS_category<-rbind(AMS_air,AMS_ocean)
AMS_category$WH<-"AMS"
AMS_category_group<-group_by(AMS_category,WH,SKU,WK,CATEGORY)
AMS_category_summ<-summarise(AMS_category_group, QTY=sum(QTY))
AMS_ctgry<-spread(AMS_category_summ,'CATEGORY','QTY',fill=0)

#ship method (Air Ocean) pivot
AMS_air_MOT<-AMS_air_supply_df
colnames(AMS_air_MOT)[4]='WK'
AMS_air_MOT <-summarise(group_by(AMS_air_MOT,SKU,WK),Air_QTY=sum(`AMS Air QTY`))

AMS_ocean_MOT<-AMS_ocean_supply_df
colnames(AMS_ocean_MOT)[4]='WK'
AMS_ocean_MOT <-summarise(group_by(AMS_ocean_MOT,SKU,WK),Ocean_QTY=sum(`AMS Ocean QTY`))

AMS_MOT<-merge(AMS_air_MOT,AMS_ocean_MOT,all = TRUE)
AMS_MOT$WH="AMS"
AMS_MOT[is.na(AMS_MOT)]=0

# Merge (category, ship method)

AMS_total=merge(AMS_ctgry,AMS_MOT,all = TRUE)
AMS_total=AMS_total[,c("WH","SKU","WK","in-transit","ship plan","Air_QTY","Ocean_QTY")]



#############################################################################################################

########## Hong Kong ##############
HK_air_supply_df<-supply_df_active[,c(1,2,11:12)]
HK_air_supply_df<-HK_air_supply_df[which(!is.na(HK_air_supply_df$`HK Air ETA WK`)),]

HK_ocean_supply_df<-supply_df_active[,c(1,2,13:14)]
HK_ocean_supply_df<-HK_ocean_supply_df[which(!is.na(HK_ocean_supply_df$`HK Ocean ETA WK`)),]



#Category (in-transit, ship plan, build plan) pivot
HK_air=HK_air_supply_df
colnames(HK_air)[3:4]=c("QTY","WK")

HK_ocean=HK_ocean_supply_df
colnames(HK_ocean)[3:4]=c("QTY","WK")

HK_category<-rbind(HK_air,HK_ocean)
HK_category$WH<-"HK"
HK_category_group<-group_by(HK_category,WH,SKU,WK,CATEGORY)
HK_category_summ<-summarise(HK_category_group, QTY=sum(QTY))
HK_ctgry<-spread(HK_category_summ,'CATEGORY','QTY',fill=0)

#ship method (Air Ocean) pivot
HK_air_MOT<-HK_air_supply_df
colnames(HK_air_MOT)[4]='WK'
HK_air_MOT <-summarise(group_by(HK_air_MOT,SKU,WK),Air_QTY=sum(`HK Air QTY`))

HK_ocean_MOT<-HK_ocean_supply_df
colnames(HK_ocean_MOT)[4]='WK'
HK_ocean_MOT <-summarise(group_by(HK_ocean_MOT,SKU,WK),Ocean_QTY=sum(`HK Ocean QTY`))

HK_MOT<-merge(HK_air_MOT,HK_ocean_MOT,all = TRUE)
HK_MOT$WH="HK"
HK_MOT[is.na(HK_MOT)]=0

# Merge (category, ship method)

HK_total=merge(HK_ctgry,HK_MOT,all = TRUE)
HK_total=HK_total[,c("WH","SKU","WK","in-transit","ship plan","build plan","Air_QTY","Ocean_QTY")]



#############################################################################################################

supply_total=rbind(MP_total,AMS_total,HK_total)

setwd("C:/Users/xiaoxiongm/Desktop/Projects/SD simulation")
xlsx::write.xlsx(demand_total,"SD_simulation_raw.xlsx",sheetName='total demand', row.names=FALSE)
xlsx::write.xlsx(supply_total,"SD_simulation_raw.xlsx",sheetName='total supply', append = TRUE,row.names=FALSE)


