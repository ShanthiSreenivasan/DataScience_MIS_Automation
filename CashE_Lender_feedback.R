rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)

library(stringr)
library(lubridate)
library(data.table)
#library(DBI)
#library(RPostgreSQL)
#library(RMySQL)
library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
#library(bit64)
library(purrr)
library(janitor)

setwd("C:\\R\\Lender Feedback")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 


if(!dir.exists(paste("./Input/",thisdate,sep="")))
{
  dir.create(paste("./Input/",thisdate,sep=""))
} 

LastMonth_API <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'LastMonth_API') %>% select(mobile_no,customer_status_name,loan_amount,Customer_requested_loan_amount) %>% mutate(`sheetname`='LastMonth_API')

CurMonth_API <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'CurMonth_API') %>% select(mobile_no,customer_status_name,loan_amount,Customer_requested_loan_amount) %>% mutate(`sheetname`='CurMonth_API')


LastMonth_Notregister <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'LastMonth_Notregister') %>% select(mobile) %>% mutate(`customer_status_name`='AIP Approved/ Interested in docs', `loan_amount`='', `Customer_requested_loan_amount`='',`sheetname`='LastMonth_Notregister')

colnames(LastMonth_Notregister)[which(names(LastMonth_Notregister) == "mobile")] <- "mobile_no"


CurMonth_Notregister <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'CurMonth_Notregister') %>% select(mobile) %>% mutate(`customer_status_name`='AIP Approved/ Interested in docs', `loan_amount`='',`Customer_requested_loan_amount`='',`sheetname`='CurMonth_Notregister')


colnames(CurMonth_Notregister)[which(names(CurMonth_Notregister) == "mobile")] <- "mobile_no"




curmonth_disbursal <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'curmonth_disbursal') %>% select(mobile_no,customer_status_name,loan_amount,Customer_requested_loan_amount) %>% mutate(`sheetname`='curmonth_disbursal')

dff1<-LastMonth_API %>% full_join(CurMonth_API)

dff2<-LastMonth_Notregister %>% full_join(CurMonth_Notregister)


dff3<-rbind(dff1,curmonth_disbursal)


CashE_lender_MIS<-rbind(dff3,dff2)

appos_map_CashE <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Cashe')

appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(name %in% c('CashE')) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(name %in% c('CashE')) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)


appos_dump<-rbind(appos_dump1,appos_dump2)

########################CashE LenderFeedBack
CashE_appos_dump <- appos_dump %>% filter(name %in% c('CashE'))



CashE_appos_dump$ph_no<-CashE_lender_MIS$mobile_no[match(CashE_appos_dump$phone_home, CashE_lender_MIS$mobile_no)]

#names(CashE_lender_MIS)

#class(CashE_lender_MIS$mobile_no)

CashE_lender_MIS$loan_amount <- as.numeric(CashE_lender_MIS$loan_amount)

CashE_lender_MIS$mobile_no <- as.numeric(CashE_lender_MIS$mobile_no)

CashE_lender_MIS$offer_application_number<-CashE_appos_dump$offer_application_number[match(CashE_lender_MIS$mobile_no, CashE_appos_dump$phone_home)]


colnames(CashE_lender_MIS)[which(names(CashE_lender_MIS) == "mobile_no")] <- "mobile_number"

CashE_lender_MIS$CURRENT_appops_status<-CashE_appos_dump$appops_status_code[match(CashE_lender_MIS$mobile_number, CashE_appos_dump$phone_home)]


CashE_lender_MIS$New_appops_status<-appos_map_CashE$New_Status[match(CashE_lender_MIS$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]



#CashE_lender_MIS$NEW_appops_description<-appos_map_CashE$New_Status[match(CashE_lender_MIS$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]

CashE_lender_MIS$NEW_appops_description<-appos_map_CashE$Status_Description[match(CashE_lender_MIS$New_appops_status, appos_map_CashE$New_Status)]

CashE_lender_MIS <-CashE_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_CRM',
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (sheetname %in% c('LastMonth_Notregister') & New_appops_status %in% c(390)) ~ '460',
    #(sheetname %in% c('CurMonth_Notregister') & customer_status_name %in% c('AIP Approved/ Interested in docs')) ~ '390',
    # 
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

CashE_lender_MIS$Remark <- ifelse(is.na(CashE_lender_MIS$Remark), CashE_lender_MIS$New_appops_status, CashE_lender_MIS$Remark)


#CashE_lender_MIS$Remark[CashE_lender_MIS$customer_status_name == "Credit Approved", is.na(CashE_lender_MIS$loan_amount), CashE_lender_MIS$New_appops_status==990] <- "Repeated_Feedback_cases"


CashE_lender_MIS$Upload_Remarks <- paste(CashE_lender_MIS$NEW_appops_description, "-", CashE_lender_MIS$loan_amount)



CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CashE_lender_MIS$New_appops_status,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, CashE_lender_MIS$New_appops_status %in% c(280,380) 

CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CashE_lender_MIS$New_appops_status,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, CashE_lender_MIS$New_appops_status %in% c(480) 

CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CashE_lender_MIS$New_appops_status,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, CashE_lender_MIS$New_appops_status %in% c(580) 

CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CashE_lender_MIS$New_appops_status,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status %in% c(650,660,670,690,700,701,780)] <- 680#, CashE_lender_MIS$New_appops_status %in% c(680) 


CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CashE_lender_MIS$New_appops_status,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status %in% c(650,660,670,690,700,701,780)] <- 680#, CashE_lender_MIS$New_appops_status %in% c(680) 

CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 390 & !is.na(CashE_lender_MIS$Application_Number) & grepl('LastMonth_Notregister',CashE_lender_MIS$sheetname,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status <= 389] <- 390

CashE_lender_MIS$Remark[CashE_lender_MIS$Remark == 390 & !is.na(CashE_lender_MIS$Application_Number) & grepl('LastMonth_Notregister',CashE_lender_MIS$sheetname,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status >= 390] <- 460

CashE_lender_MIS$Remark[grepl('LastMonth_Notregister',CashE_lender_MIS$sheetname,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status == 390 & !is.na(CashE_lender_MIS$Application_Number) & CashE_lender_MIS$New_appops_status == 390] <- 460

CashE_lender_MIS$Upload_Remarks[grepl('LastMonth_Notregister',CashE_lender_MIS$sheetname,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status == 390 & !is.na(CashE_lender_MIS$Application_Number) & CashE_lender_MIS$New_appops_status == 390] <- 'LastMonth_Notregister- AIP Approved/ Interested in docs- Docs NI'  

CashE_lender_MIS$NEW_appops_description[grepl('LastMonth_Notregister',CashE_lender_MIS$sheetname,ignore.case=TRUE) & CashE_lender_MIS$CURRENT_appops_status == 390 & !is.na(CashE_lender_MIS$Application_Number) & CashE_lender_MIS$New_appops_status == 390] <- 'Docs NI'  


colnames(CashE_lender_MIS)[which(names(CashE_lender_MIS) == "offer_application_number")] <- "Application_Number"


CashE_lender_MIS<-CashE_lender_MIS %>% mutate('Lender'="CashE")
write.xlsx(CashE_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_CashE_FB_File.xlsx")


CashE_upload_1<-CashE_lender_MIS %>% filter(NEW_appops_description %in% c("Loan Disbursed"),!is.na(Application_Number),!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'))

CashE_upload_1<- CashE_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=CashE_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=CashE_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=CashE_upload_1$loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


CashE_upload_2<-CashE_lender_MIS %>% filter(!NEW_appops_description %in% c("Loan Disbursed"),!is.na(Application_Number),!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'))

CashE_upload_2<- CashE_upload_2 %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not_in_CRM","Stuck_cases")) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=CashE_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=CashE_upload_2$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

#names(CashE_lender_MIS)

#write.xlsx(appos_dump, file = "C:\\R\\Lender Feedback\\Output\\CashE_appos_dump.xlsx")

#write.xlsx(CashE_upload_2, file = "C:\\R\\Lender Feedback\\Output\\CashE_upload_2.xlsx")


#################################
