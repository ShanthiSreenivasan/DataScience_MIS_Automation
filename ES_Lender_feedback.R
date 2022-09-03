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
library(DBI)
library(RPostgreSQL)
library(RMySQL)
library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)
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

ES_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\ES.xlsx") %>% select(mobile_number,status,subs_status,rejectreasonsgroup,app_download_flag,first_disb_loan_amt)

names(ES_lender_MIS)
appos_map_ES <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'EarlySalary')
ES_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Early',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################ES LenderFeedBack


colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "subs_status")] <- "customer_status_name"

ES_lender_MIS$mobile_number <- as.numeric(ES_lender_MIS$mobile_number)


ES_lender_MIS$offer_application_number<-ES_appos_dump$offer_application_number[match(ES_lender_MIS$mobile_number, ES_appos_dump$phone_home)]



ES_lender_MIS$CURRENT_appops_status<-ES_appos_dump$appops_status_code[match(ES_lender_MIS$mobile_number, ES_appos_dump$phone_home)]


ES_lender_MIS$New_appops_status<-appos_map_ES$New_Status[match(ES_lender_MIS$customer_status_name, appos_map_ES$CM_Status)]

ES_lender_MIS$NEW_appops_description<-appos_map_ES$Status_Description[match(ES_lender_MIS$customer_status_name, appos_map_ES$CM_Status)]

names(ES_lender_MIS)


ES_lender_MIS <-ES_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(offer_application_number) ~ 'Not in CRM',
    app_download_flag==0 & New_appops_status<=390 ~ '390',
    app_download_flag==1 & New_appops_status<=399 ~ '399',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
  ))

ES_lender_MIS$Remark <- ifelse(is.na(ES_lender_MIS$Remark), ES_lender_MIS$New_appops_status, ES_lender_MIS$Remark)


ES_Not_in_CRM_cases <- ES_appos_dump %>% filter(is.na(phone_home))


ES_lender_MIS$Upload_Remarks <- paste(ES_lender_MIS$customer_status_name,'-',ES_lender_MIS$rejectreasonsgroup,'-',ES_lender_MIS$NEW_appops_description)

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "offer_application_number")] <- "Application_Number"

write.xlsx(ES_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_ES_FB_File.xlsx")

ES_upload<-ES_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

ES_upload<- ES_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=ES_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=ES_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(ES_upload, file = "C:\\R\\Lender Feedback\\Output\\ES_upload.xlsx")

#################################
