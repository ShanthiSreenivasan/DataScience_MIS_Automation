
rm(list = ls())

library(data.table)
library(readxl)
library(mailR)
library(xtable)
library(dplyr)
library(stringi)
library(xlsx)
library(dtplyr)
library(xlsx)
library(lubridate)
library(magrittr)

today=format(Sys.Date(), format="%d-%B-%Y")

x<-"No New leads"

#dfAxis<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\Axis.xlsx",sheet = 'Customer_Details')



dfAxis1<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\AXIS1.xlsx",sheet = 'Customer_Details') %>% distinct()

dfAxis2<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\AXIS2.xlsx",sheet = 'Customer_Details')%>% distinct()

 
 
#dfAxis3<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\Axis3.xlsx")%>% distinct()



Axis1<-"AXIS1.xlsx"
Axis2<-"AXIS2.xlsx"
# 
# Axis3<-"Axis3.xlsx"

if(nrow(dfAxis1>=0) | nrow(dfAxis2>=0)){
 
     dfAxis<-rbind(dfAxis1,dfAxis2) %>% select(-c(S.NO)) %>% as.data.frame()
     df1<-xlsx::write.xlsx(as.data.frame(dfAxis), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Axis_",today,".xlsx"), password = "Cmaxis!123", row.names = FALSE)
     #dfA3<-xlsx::write.xlsx(as.data.frame(Axis3), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Axis_PL_Interested_Customers_",today,".xlsx"), password = "Cmaxis!123", row.names = FALSE)
     
     #list_of_datasets <- list("New leads" = dfAxis, "Interested Customers" = dfA3)
     #write.xlsx(list_of_datasets, paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Axis_",today,".xlsx"), password = "Cmaxis!123", row.names = FALSE)
     
 }else if(!file.exists(AXIS2)){
   df1<-xlsx::write.xlsx(as.data.frame(dfAxis), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Axis_",today,".xlsx"), password = "Cmaxis!123", row.names = FALSE)
   
 }
############################
dfAxis<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\AXIS.xlsx",sheet = 'Customer_Details') %>% distinct()

if(nrow(dfAxis>=0)){
  
  df1<-xlsx::write.xlsx(as.data.frame(dfAxis), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Axis_",today,".xlsx"), password = "Cmaxis!123", row.names = FALSE)
  
}else{
  print(x)
}

############################
dfbaj<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\BAJ.xlsx",sheet = 'Customer_Details') %>% distinct()

if(nrow(dfbaj>=0)){
  
  dfb<-xlsx::write.xlsx(as.data.frame(dfbaj), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Bajaj_",today,".xlsx"), password = "Cmbajaj!123", row.names = FALSE)
  
}else{
  print(x)
}
###########3
dfPNB<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\PNB.xlsx")%>% distinct()

#dfPNB<-dfPNB.drop(dfPNB.index[[1,2]])

if(nrow(dfPNB>=0)){
  
  df8<-xlsx::write.xlsx(as.data.frame(dfPNB), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\PNB_",today,".xlsx"), password = "Cmpnb!123", row.names = FALSE)
  
  
}else{
  print(x)
}

############################

dfMuthoot1<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\MUTH1.xlsx")%>% distinct()

dfMuthoot2<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\MUTH2.xlsx")%>% distinct()

Muthoot1<-"MUTH1.xlsx"
Muthoot2<-"MUTH2.xlsx"

if(nrow(dfMuthoot1>=0) & nrow(dfMuthoot2>=0)){
  
  dfMuthoot<-rbind(dfMuthoot1,dfMuthoot2) %>% select(-c(S.NO))
  df7<-xlsx::write.xlsx(as.data.frame(dfMuthoot), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Muthoot_",today,".xlsx"), password = "Cmmuthoot!123", row.names = FALSE)
  
}else if(!file.exists(MUTH2)){
  df11<-xlsx::write.xlsx(as.data.frame(dfMuthoot1), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Muthoot_",today,".xlsx"), password = "Cmmuthoot!123", row.names = FALSE)
}


############################
dfMuthoot<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\MUTH.xlsx",sheet = 'Customer_Details') %>% distinct()

if(nrow(dfMuthoot>=0)){
  
  df1<-xlsx::write.xlsx(as.data.frame(dfMuthoot), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Muthoot_",today,".xlsx"), password = "Cmmuthoot!123", row.names = FALSE)
  
}else{
  print(x)
}


dfIS<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\INDS.xlsx") %>% distinct()


IS<-"INDS.xlsx"

if(nrow(dfIS>=0)){
  
  df5<-xlsx::write.xlsx(as.data.frame(dfIS), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\IndiaShelter_",today,".xlsx"), password = "Cmindiashelter!123", row.names = FALSE)
  
}else{
  print(x)
}


dfCiti<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\CITI.xlsx")%>% distinct()

Citi<-"CITI.xlsx"

if(nrow(dfCiti>=0)){
  
  df3<-xlsx::write.xlsx(as.data.frame(dfCiti), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\CITI_",today,".xlsx"), password = "CM07", row.names = FALSE)
  
}else{
  print(x)
}


dfAye<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\AYE.xlsx")%>% distinct()


Aye<-"AYE.xlsx"

if(nrow(dfAye>=0)){
  
  df2<-xlsx::write.xlsx(as.data.frame(dfAye), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\AyeFinance_",today,".xlsx"), password = "Cmaye!123", row.names = FALSE)
  
}else{
  print(x)
}


dfIndi<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\INDIF.xlsx")%>% distinct()


#Indi<-"Indifi.xlsx"

if(nrow(dfIndi>=0)){
  
  Indif<-xlsx::write.xlsx(as.data.frame(dfIndi), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Indifi_",today,".xlsx"), password = "Cmindifi!123", row.names = FALSE)
  
}else{
  print(x)
}


dfIncre<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\INC.xlsx")%>% distinct()



if(nrow(dfIncre>=0)){
  
  df2<-xlsx::write.xlsx(as.data.frame(dfIncre), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Incred Omega_",today,".xlsx"), password = "Cmincred!123", row.names = FALSE)
  
}else{
  print(x)
}

dfmcap<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\MCAP.xlsx")%>% distinct()


#Indi<-"Indifi.xlsx"

if(nrow(dfmcap>=0)){
  
  mcap<-xlsx::write.xlsx(as.data.frame(dfmcap), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\M Capital_",today,".xlsx"), password = "Cmmcapital!123", row.names = FALSE)
  
}else{
  print(x)
}

dfindg<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\INDG.xlsx")%>% distinct()


#Indi<-"Indifi.xlsx"

if(nrow(dfindg>=0)){
  
  ing<-xlsx::write.xlsx(as.data.frame(dfindg), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\India Gold_",today,".xlsx"), password = "Cmindiagold!123", row.names = FALSE)
  
}else{
  print(x)
}

rup<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\RUP.xlsx")%>% distinct()



if(nrow(rup>=0)){
  
  rup_df<-xlsx::write.xlsx(as.data.frame(rup), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Rupeek_",today,".xlsx"), password = "Cmrupeek!123", row.names = FALSE)
  
}else{
  print(x)
}

dfrupt<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\RUPT.xlsx")%>% distinct()



if(nrow(dfrupt>=0)){
  
  rupt<-xlsx::write.xlsx(as.data.frame(dfrupt), paste0("C:\\R\\S2L_Manual_Lenders\\Output\\RUPTOK_",today,".xlsx"), password = "Cmruptok!123", row.names = FALSE)
  
}else{
  print(x)
}


###########################S2L upload file 

dfdm<-read_excel("C:\\R\\S2L_Manual_Lenders\\Input\\DM.xlsx")

dfdm<-dfdm %>% filter(!Application_Number=='') %>% mutate(App_Ops_Status="Application forwarded to Lender") %>% select(Application_Number,App_Ops_Status) 
write.xlsx(dfdm, file = paste0("C:\\R\\S2L_Manual_Lenders\\Output\\Upload 1 ref ", today, '.xlsx'))
