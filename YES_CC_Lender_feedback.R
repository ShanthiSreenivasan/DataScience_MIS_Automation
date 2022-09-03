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

YES_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\YES_CC.xlsx", sheet = "Sheet1")# %>% select(customer_contact,status,Dip_Status,DIP.REJECT.REASON)



appos_map_YES <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Yesback cc')
YES_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Yes Bank',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

names(YES_lender_MIS)
########################YES LenderFeedBack

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "customer_contact")] <- "mobile_number"

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "status")] <- "customer_status_name"

YES_lender_MIS$mobile_number <- as.numeric(YES_lender_MIS$mobile_number)


YES_lender_MIS$offer_application_number<-YES_appos_dump$offer_application_number[match(YES_lender_MIS$mobile_number, YES_appos_dump$phone_home)]



YES_lender_MIS$CURRENT_appops_status<-YES_appos_dump$appops_status_code[match(YES_lender_MIS$mobile_number, YES_appos_dump$phone_home)]


YES_lender_MIS$New_appops_status<-appos_map_YES$New_Status[match(YES_lender_MIS$customer_status_name, appos_map_YES$CM_Status)]

YES_lender_MIS$NEW_appops_description<-appos_map_YES$Status_Description[match(YES_lender_MIS$customer_status_name, appos_map_YES$CM_Status)]

names(YES_lender_MIS)
YES_lender_MIS <-YES_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(offer_application_number) ~ 'Not in CRM',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

YES_lender_MIS$Remark <- ifelse(is.na(YES_lender_MIS$Remark), YES_lender_MIS$New_appops_status, YES_lender_MIS$Remark)

#YES_lender_MIS<-YES_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
YES_lender_MIS$Upload_Remarks <- paste(YES_lender_MIS$NEW_appops_description,'-',YES_lender_MIS$customer_status_name)

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "offer_application_number")] <- "Application_Number"

YES_upload<- YES_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not in CRM")) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=YES_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=YES_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(YES_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_YES_FB_File.xlsx")

write.xlsx(YES_upload, file = "C:\\R\\Lender Feedback\\Output\\YES_upload.xlsx")

#################################
