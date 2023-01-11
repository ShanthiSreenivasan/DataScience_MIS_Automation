rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)

library(stringr)
library(lubridate)
library(openxlsx)
library(janitor)
library(data.table)
library(readxl)
library(mailR)
library(xtable)
library(dplyr)
library(plyr)

setwd("C:\\R\\SCUF")


today_dat <-format(Sys.Date(),'%Y-%m-%d')


thisDate = Sys.Date()
if(!dir.exists(paste(".\\Input\\",thisDate,sep="")))
{
  dir.create(paste(".\\Input\\",thisDate,sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,sep=""))
} 



ASSISTED_dump <- read.xlsx("C:\\R\\SCUF\\Input\\SCUF Working sheet.xlsx",sheet='Assisted')

ASSISTED_dump$Lead.id <- as.numeric(ASSISTED_dump$Lead.id)

ASSISTED_dump$Status <- as.numeric(ASSISTED_dump$Status)

ASSISTED_dump$slug <- as.character(ASSISTED_dump$slug)


#names(b_df)


a_df<-ASSISTED_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks) %>% dplyr::summarise(Total = n_distinct(Lead.id, na.rm = TRUE))# %>% adorn_totals("row")

a_df<-spread(a_df,key=Remarks, value = Total)


a_df<-a_df# %>% adorn_totals("col")

MTD_ASSISTED_pending<-ASSISTED_dump %>% filter(is.na(ASSISTED_dump$Remarks))

a_df$Pending<-ifelse((length(MTD_ASSISTED_pending)==0), 0, nrow(MTD_ASSISTED_pending))


a_df<-a_df %>% mutate(`Inflow`=sum(Approved+Error+Rejected+Pending),`Processed`=sum(Approved+Error+Rejected), `Approved %`=scales::percent((Approved/Inflow), accuracy = 0.01), `Rejected %`=scales::percent((Rejected/Inflow), accuracy = 0.01),
                      `Error %`=scales::percent((Error/Inflow), accuracy = 0.01),`Source`='Assisted')

LIVE_dump <- read.xlsx("C:\\R\\SCUF\\Input\\SCUF Working sheet.xlsx",sheet='Live')

LIVE_dump$Lead.id <- as.numeric(LIVE_dump$Lead.id)

LIVE_dump$Status <- as.numeric(LIVE_dump$Status)

LIVE_dump$slug <- as.character(LIVE_dump$slug)

#names(b_df)


b_df<-LIVE_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks) %>% dplyr::summarise(Total = n_distinct(Lead.id, na.rm = TRUE))# %>% adorn_totals("row")

b_df<-spread(b_df,key=Remarks, value = Total)


b_df<-b_df


MTD_LIVE_pending<-LIVE_dump %>% filter(is.na(LIVE_dump$Remarks))

b_df$Pending<-ifelse((length(MTD_LIVE_pending)==0), 0, nrow(MTD_LIVE_pending))


b_df<-b_df %>% mutate(`Inflow`=sum(Approved+Error+Rejected+Pending),`Processed`=sum(Approved+Error+Rejected), `Approved %`=scales::percent((Approved/Inflow), accuracy = 0.01), `Rejected %`=scales::percent((Rejected/Inflow), accuracy = 0.01),
                      `Error %`=scales::percent((Error/Inflow), accuracy = 0.01),`Source`='STP/PA')

#final_DF<-rbind(a_df,b_df)
#names(final_DF)
final_DF<-rbind(a_df,b_df) 

MTD_final_DF_FINAL<-select(final_DF,10,5,6,4,1,7,3,8,2,9)

MTD_final_DF_FINAL<-MTD_final_DF_FINAL %>% adorn_totals("row")
DTD_ASSISTED_dump <- read.xlsx("C:\\R\\SCUF\\Input\\SCUF Working sheet.xlsx",sheet='Assisted') %>% filter(Date=='TODAY')

DTD_ASSISTED_dump$Lead.id <- as.numeric(DTD_ASSISTED_dump$Lead.id)

DTD_ASSISTED_dump$Status <- as.numeric(DTD_ASSISTED_dump$Status)

DTD_ASSISTED_dump$slug <- as.character(DTD_ASSISTED_dump$slug)

DTD_a_df<-DTD_ASSISTED_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks) %>% dplyr::summarise(Total = n_distinct(Lead.id, na.rm = TRUE))# %>% adorn_totals("row")

DTD_a_df<-spread(DTD_a_df,key=Remarks, value = Total)


DTD_a_df<-DTD_a_df# %>% adorn_totals("col")



DTD_ASSISTED_pending<-DTD_ASSISTED_dump %>% filter(is.na(DTD_ASSISTED_dump$Remarks))

DTD_a_df$Pending<-ifelse((length(DTD_ASSISTED_pending)==0), 0, nrow(DTD_ASSISTED_pending))

DTD_a_df<-DTD_a_df %>% mutate(`Inflow`=sum(Approved+Error+Rejected+Pending),`Processed`=sum(Approved+Error+Rejected), `Approved %`=scales::percent((Approved/Inflow), accuracy = 0.01), `Rejected %`=scales::percent((Rejected/Inflow), accuracy = 0.01),
                              `Error %`=scales::percent((Error/Inflow), accuracy = 0.01),`Source`='Assisted')

DTD_LIVE_dump <- read.xlsx("C:\\R\\SCUF\\Input\\SCUF Working sheet.xlsx",sheet='Live') %>% filter(Date=='TODAY')

DTD_LIVE_dump$Lead.id <- as.numeric(DTD_LIVE_dump$Lead.id)

DTD_LIVE_dump$Status <- as.numeric(DTD_LIVE_dump$Status)

DTD_LIVE_dump$slug <- as.character(DTD_LIVE_dump$slug)

#names(DTD_b_df)


DTD_b_df<-DTD_LIVE_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks) %>% dplyr::summarise(Total = n_distinct(Lead.id, na.rm = TRUE))# %>% adorn_totals("row")

DTD_b_df<-spread(DTD_b_df,key=Remarks, value = Total)


DTD_b_df<-DTD_b_df# %>% adorn_totals("col")


DTD_LIVE_pending<-DTD_LIVE_dump %>% filter(is.na(DTD_LIVE_dump$Remarks))

DTD_b_df$Pending<-ifelse((length(DTD_LIVE_pending)==0), 0, nrow(DTD_LIVE_pending))

DTD_b_df<-DTD_b_df %>% mutate(`Inflow`=sum(Approved+Error+Rejected+Pending),`Processed`=sum(Approved+Error+Rejected), `Approved %`=scales::percent((Approved/Inflow), accuracy = 0.01), `Rejected %`=scales::percent((Rejected/Inflow), accuracy = 0.01),
                              `Error %`=scales::percent((Error/Inflow), accuracy = 0.01),`Source`='STP/PA')

DTD_final_DF<-rbind(DTD_a_df,DTD_b_df)

DTD_final_DF_FINAL<-select(DTD_final_DF,10,5,6,4,1,7,3,8,2,9)


DTD_final_DF_FINAL<-DTD_final_DF_FINAL %>% adorn_totals("row")

DTD_REJ_df1<-DTD_ASSISTED_dump %>% filter(grepl('Rejected',Remarks,ignore.case=TRUE)) %>% group_by(Reject.Reason) %>% dplyr::summarise(`ASSISTED` = n_distinct(Lead.id, na.rm = TRUE))


DTD_REJ_df2<- DTD_LIVE_dump %>% filter(grepl('Rejected',Remarks,ignore.case=TRUE)) %>% group_by(Reject.Reason) %>% dplyr::summarise(`STP/PA` = n_distinct(Lead.id, na.rm = TRUE))
DTD_REJ_df= DTD_REJ_df1 %>% full_join(DTD_REJ_df2,by="Reject.Reason")


colnames(DTD_REJ_df)[which(names(DTD_REJ_df) == "Total.x")] <- "ASSISTED"
colnames(DTD_REJ_df)[which(names(DTD_REJ_df) == "Total.y")] <- "STP/PA"

DTD_REJ_df <- replace(DTD_REJ_df, is.na(DTD_REJ_df), 0)

DTD_REJ_df<-DTD_REJ_df %>% adorn_totals("col") 

DTD_REJ_df<-DTD_REJ_df %>% adorn_totals("row")

DTD_ERR_df1<-DTD_ASSISTED_dump %>% filter(grepl('ERROR',Remarks,ignore.case=TRUE)) %>% group_by(Reject.Reason) %>% dplyr::summarise(`ASSISTED` = n_distinct(Lead.id, na.rm = TRUE))

DTD_ERR_df2<- DTD_LIVE_dump %>% filter(grepl('ERROR',Remarks,ignore.case=TRUE)) %>% group_by(Reject.Reason) %>% dplyr::summarise(`STP/PA` = n_distinct(Lead.id, na.rm = TRUE))
DTD_ERR_df= DTD_ERR_df1 %>% full_join(DTD_ERR_df2,by="Reject.Reason")


colnames(DTD_ERR_df)[which(names(DTD_ERR_df) == "Total.x")] <- "ASSISTED"
colnames(DTD_ERR_df)[which(names(DTD_ERR_df) == "Total.y")] <- "STP/PA"
colnames(DTD_ERR_df)[which(names(DTD_ERR_df) == "Reject.Reason")] <- "Error Reason"

DTD_ERR_df <- replace(DTD_ERR_df, is.na(DTD_ERR_df), 0)

DTD_ERR_df<-DTD_ERR_df %>% adorn_totals("col") 

DTD_ERR_df<-DTD_ERR_df %>% adorn_totals("row")


DTD_df1<-DTD_ASSISTED_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks,slug) %>% dplyr::summarise(`Total` = n_distinct(Lead.id, na.rm = TRUE))

DTD_df1_FINAL<-spread(DTD_df1,key=Remarks, value = Total)


DTD_df1_FINAL<-DTD_df1_FINAL %>% adorn_totals("col") 
DTD_df1_FINAL <- replace(DTD_df1_FINAL, is.na(DTD_df1_FINAL), 0)

colnames(DTD_df1_FINAL)[which(names(DTD_df1_FINAL) == "Total")] <- "Processed"

DTD_df1_FINAL<-DTD_df1_FINAL %>% adorn_totals("row") 

DTD_df1_FINAL<-select(DTD_df1_FINAL,1,5,2,3,4)
#names(DTD_df_FINAL)
MTD_df1<-ASSISTED_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks,slug) %>% dplyr::summarise(`Total` = n_distinct(Lead.id, na.rm = TRUE))

MTD_df1_FINAL<-spread(MTD_df1,key=Remarks, value = Total)


MTD_df1_FINAL<-MTD_df1_FINAL %>% adorn_totals("col") 
MTD_df1_FINAL <- replace(MTD_df1_FINAL, is.na(MTD_df1_FINAL), 0)

colnames(MTD_df1_FINAL)[which(names(MTD_df1_FINAL) == "Total")] <- "Processed"

MTD_df1_FINAL<-MTD_df1_FINAL %>% adorn_totals("row") 

MTD_df1_FINAL<-select(MTD_df1_FINAL,1,5,2,3,4)




DTD_df2<-DTD_LIVE_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks,slug) %>% dplyr::summarise(`Total` = n_distinct(Lead.id, na.rm = TRUE))

DTD_df2_FINAL<-spread(DTD_df2,key=Remarks, value = Total)


DTD_df2_FINAL<-DTD_df2_FINAL %>% adorn_totals("col") 
DTD_df2_FINAL <- replace(DTD_df2_FINAL, is.na(DTD_df2_FINAL), 0)

colnames(DTD_df2_FINAL)[which(names(DTD_df2_FINAL) == "Total")] <- "Processed"

DTD_df2_FINAL<-DTD_df2_FINAL %>% adorn_totals("row") 

DTD_df2_FINAL<-select(DTD_df2_FINAL,1,5,2,3,4)
#names(DTD_df_FINAL)
MTD_df2<-LIVE_dump %>% filter(!is.na(Remarks)) %>% group_by(Remarks,slug) %>% dplyr::summarise(`Total` = n_distinct(Lead.id, na.rm = TRUE))

MTD_df2_FINAL<-spread(MTD_df2,key=Remarks, value = Total)


MTD_df2_FINAL<-MTD_df2_FINAL %>% adorn_totals("col") 
MTD_df2_FINAL <- replace(MTD_df2_FINAL, is.na(MTD_df2_FINAL), 0)

colnames(MTD_df2_FINAL)[which(names(MTD_df2_FINAL) == "Total")] <- "Processed"

MTD_df2_FINAL<-MTD_df2_FINAL %>% adorn_totals("row") 

MTD_df2_FINAL<-select(MTD_df2_FINAL,1,5,2,3,4)

today=format(Sys.Date()-2, format="%d-%B-%Y")


filename<-("C:\\R\\SCUF\\Input\\SCUF Working sheet.xlsx")

myMessage = paste0("SCUF EOD MIS- ",today,sep='')
sender <- "Credit-Ops@creditmantri.com"
# mail_LNT <-"shanthi.s@creditmantri.com"
# cclist <-"shanthi.s@creditmantri.com"
mail_LNT <- c("rakshith.thangaraj@creditmantri.com","sankarnarayanan@creditmantri.com","balamurugan.r@creditmantri.com") 
cclist <-c("Credit-Ops@creditmantri.com")

msg = '
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <style>
    table {font-family:  Verdana, Geneva, sans-serif; font-size: 11px;
    width = 100%; border: 1px solid black; border-collapse: collapse;
    text-align: center; padding: 5px;}
    th {height = 12px;background-color: #4CAF50;color: white;}
    td {background-color: #FFF;}
    </style>
  </head>
  <body>
    
    <h3> SCUF Document Validation - DTD </h3>
    <p> ${print(xtable(DTD_final_DF_FINAL, digits = 0), type = \'html\')} </p><br>
    <h3> SCUF Document Validation - MTD </h3>
    <p> ${print(xtable(MTD_final_DF_FINAL, digits = 0), type = \'html\')} </p><br>
    <h3> Reject Reason DTD Summary </h3>
    <p> ${print(xtable(DTD_REJ_df, digits = 0), type = \'html\')} </p><br>
    <h3> ERROR Reason DTD Summary(Tech Query - Tracker) </h3>
    <p> ${print(xtable(DTD_ERR_df, digits = 0), type = \'html\')} </p><br>
    <h3> Assisted - SCUF (50K-1L )Validation - DTD </h3>
    <p> ${print(xtable(DTD_df1_FINAL, digits = 0), type = \'html\')} </p><br>
    <h3> Assisted - SCUF (50K-1L )Validation - MTD </h3>
    <p> ${print(xtable(MTD_df1_FINAL, digits = 0), type = \'html\')} </p><br>
    
    <h3> STP/PA - SCUF (50K-1L )Validation - DTD </h3>
    <p> ${print(xtable(DTD_df2_FINAL, digits = 0), type = \'html\')} </p><br>
    <h3> STP/PA - SCUF (50K-1L )Validation - MTD </h3>
    <p> ${print(xtable(MTD_df2_FINAL, digits = 0), type = \'html\')} </p><br>
    
</body>
</html>'


if(file.exists(filename)){
  email <- send.mail(from = sender,
                     to = mail_LNT,
                     cc = c(cclist),
                     subject=myMessage,
                     html = TRUE,
                     inline = T,
                     body = str_interp(msg),
                     smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                                 user.name = "AKIAXJKXCWJGAXCHHIGY",
                                 passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}




