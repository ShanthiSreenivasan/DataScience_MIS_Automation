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

CITI_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CITI_CC.xlsx") %>% select(Creative,V_APPLICATION_ID,final_status,reject.reasons)

view(CITI_lender_MIS)
appos_map_CITI <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'CITICC')
CITI_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('CITI',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)
########################CITI LenderFeedBack

colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "Creative")] <- "leadid"

colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "final_status")] <- "customer_status_name"


CITI_lender_MIS$leadid <- as.numeric(CITI_lender_MIS$leadid)


CITI_lender_MIS$offer_application_number<-CITI_appos_dump$offer_application_number[match(CITI_lender_MIS$leadid, CITI_appos_dump$phone_home)]



CITI_lender_MIS$CURRENT_appops_status<-CITI_appos_dump$appops_status_code[match(CITI_lender_MIS$leadid, CITI_appos_dump$phone_home)]


CITI_lender_MIS$New_appops_status<-appos_map_CITI$New_Status[match(CITI_lender_MIS$customer_status_name, appos_map_CITI$CM_Status)]

CITI_lender_MIS$NEW_appops_description<-appos_map_CITI$Status_Description[match(CITI_lender_MIS$customer_status_name, appos_map_CITI$CM_Status)]


CITI_lender_MIS <-CITI_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(leadid) ~ 'Not in CRM',
    is.na(offer_application_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

CITI_lender_MIS$Remark <- ifelse(is.na(CITI_lender_MIS$Remark), CITI_lender_MIS$New_appops_status, CITI_lender_MIS$Remark)

#CITI_lender_MIS<-CITI_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
CITI_lender_MIS$Upload_Remarks <- paste(CITI_lender_MIS$NEW_appops_description,'-',CITI_lender_MIS$customer_status_name)

colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "offer_application_number")] <- "Application_Number"

CITI_upload<- CITI_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not in CRM")) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=CITI_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=CITI_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(CITI_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_CITI_FB_File.xlsx")

write.xlsx(CITI_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\appos_map_CITI.xlsx")

#################################
