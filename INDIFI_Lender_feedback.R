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
library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
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

INDIFI_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\INDIFI.xlsx") %>% select(phone,state,applied_amount,rejection_reason)




names(INDIFI_lender_MIS)
appos_map_INDIFI <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'INDIFI')
INDIFI_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('INDIFI',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

########################INDIFI LenderFeedBack

INDIFI_lender_MIS$phone <- as.numeric(INDIFI_lender_MIS$phone)

colnames(INDIFI_lender_MIS)[which(names(INDIFI_lender_MIS) == "phone")] <- "mobile_number"

colnames(INDIFI_lender_MIS)[which(names(INDIFI_lender_MIS) == "state")] <- "customer_status_name"

#INDIFI_appos_dump$ph_no<-INDIFI_lender_MIS$mobile_number[match(INDIFI_appos_dump$phone_home, INDIFI_lender_MIS$mobile_number)]

INDIFI_lender_MIS$offer_application_number<-INDIFI_appos_dump$offer_application_number[match(INDIFI_lender_MIS$mobile_number, INDIFI_appos_dump$phone_home)]



INDIFI_lender_MIS$CURRENT_appops_status<-INDIFI_appos_dump$appops_status_code[match(INDIFI_lender_MIS$mobile_number, INDIFI_appos_dump$phone_home)]

INDIFI_lender_MIS$New_appops_status<-appos_map_INDIFI$New_Status[match(INDIFI_lender_MIS$customer_status_name, appos_map_INDIFI$CM_Status)]

INDIFI_lender_MIS$NEW_appops_description<-appos_map_INDIFI$Status_Description[match(INDIFI_lender_MIS$customer_status_name, appos_map_INDIFI$CM_Status)]


INDIFI_lender_MIS <-INDIFI_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    is.na(offer_application_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

INDIFI_lender_MIS$Remark <- ifelse(is.na(INDIFI_lender_MIS$Remark), INDIFI_lender_MIS$New_appops_status, INDIFI_lender_MIS$Remark)

INDIFI_lender_MIS$Upload_Remarks <- paste(INDIFI_lender_MIS$customer_status_name,"-",INDIFI_lender_MIS$rejection_reason)


colnames(INDIFI_lender_MIS)[which(names(INDIFI_lender_MIS) == "offer_application_number")] <- "Application_Number"

INDIFI_lender_MIS<-INDIFI_lender_MIS %>% mutate('Lender'="INDIFI")


write.xlsx(INDIFI_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_INDIFI_FB_File.xlsx")

INDIFI_upload_1<-INDIFI_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))





INDIFI_upload_1<- INDIFI_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=INDIFI_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=INDIFI_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(INDIFI_upload_1, file = "C:\\R\\Lender Feedback\\Output\\INDIFI_upload.xlsx")

