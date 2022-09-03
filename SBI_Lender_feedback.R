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

SBI_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\SBI.xlsx", sheet = "Sheet1")# %>% select(GEMID1,APPLICATION_STAGE,STATUS,Decline.Description)




#names(SBI_lender_MIS)
appos_map_SBI <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'SBI cc')
SBI_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('SBI',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

########################SBI LenderFeedBack

SBI_lender_MIS$GEMID1 <- as.numeric(SBI_lender_MIS$GEMID1)

colnames(SBI_lender_MIS)[which(names(SBI_lender_MIS) == "GEMID1")] <- "leadid"

colnames(SBI_lender_MIS)[which(names(SBI_lender_MIS) == "STATUS")] <- "customer_status_name"

SBI_lender_MIS$phone_home<-SBI_appos_dump$phone_home[match(SBI_lender_MIS$leadid,SBI_appos_dump$leadid)]

SBI_lender_MIS$offer_application_number<-SBI_appos_dump$offer_application_number[match(SBI_lender_MIS$leadid, SBI_appos_dump$leadid)]



SBI_lender_MIS$CURRENT_appops_status<-SBI_appos_dump$appops_status_code[match(SBI_lender_MIS$leadid, SBI_appos_dump$leadid)]

SBI_lender_MIS$New_appops_status<-appos_map_SBI$New_Status[match(SBI_lender_MIS$customer_status_name, appos_map_SBI$CM_Status)]

SBI_lender_MIS$NEW_appops_description<-appos_map_SBI$Status_Description[match(SBI_lender_MIS$customer_status_name, appos_map_SBI$CM_Status)]


SBI_lender_MIS <-SBI_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    is.na(offer_application_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

SBI_lender_MIS$Remark <- ifelse(is.na(SBI_lender_MIS$Remark), SBI_lender_MIS$New_appops_status, SBI_lender_MIS$Remark)

SBI_lender_MIS$Upload_Remarks <- paste(SBI_lender_MIS$APPLICATION_STAGE,"-",SBI_lender_MIS$Decline.Description)


colnames(SBI_lender_MIS)[which(names(SBI_lender_MIS) == "offer_application_number")] <- "Application_Number"

SBI_lender_MIS<-SBI_lender_MIS %>% mutate('Lender'="SBI")

write.xlsx(SBI_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SBI_FB_File.xlsx")

SBI_upload_1<-SBI_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))





SBI_upload_1<- SBI_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=SBI_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=SBI_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(SBI_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SBI_upload.xlsx")

