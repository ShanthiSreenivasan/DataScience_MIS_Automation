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

Protium_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Protium.xlsx") %>% select(appl1_mobile,lead_id,disposition,closure_reason)

names(Protium_lender_MIS)
appos_map_Protium <- read_xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Protium')
Protium_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Protium',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################Protium LenderFeedBack

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "appl1_mobile")] <- "mobile_number"

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "disposition")] <- "customer_status_name"

Protium_lender_MIS$mobile_number <- as.numeric(Protium_lender_MIS$mobile_number)


Protium_lender_MIS$offer_application_number<-Protium_appos_dump$offer_application_number[match(Protium_lender_MIS$mobile_number, Protium_appos_dump$phone_home)]



Protium_lender_MIS$CURRENT_appops_status<-Protium_appos_dump$appops_status_code[match(Protium_lender_MIS$mobile_number, Protium_appos_dump$phone_home)]


Protium_lender_MIS$New_appops_status<-appos_map_Protium$New_Status[match(Protium_lender_MIS$customer_status_name, appos_map_Protium$CM_Status)]

Protium_lender_MIS$NEW_appops_description<-appos_map_Protium$Status_Description[match(Protium_lender_MIS$customer_status_name, appos_map_Protium$CM_Status)]


Protium_lender_MIS <-Protium_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

Protium_lender_MIS$Remark <- ifelse(is.na(Protium_lender_MIS$Remark), Protium_lender_MIS$New_appops_status, Protium_lender_MIS$Remark)

Protium_lender_MIS$Upload_Remarks <- paste(Protium_lender_MIS$Status,'-',Protium_lender_MIS$customer_status_name)

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "offer_application_number")] <- "Application_Number"


Protium_lender_MIS<-Protium_lender_MIS %>% mutate('Lender'="Protium")

write.xlsx(Protium_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Protium_FB_File.xlsx")

Protium_upload<-Protium_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

Protium_upload<- Protium_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Protium_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Protium_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(Protium_upload, file = "C:\\R\\Lender Feedback\\Output\\Protium_upload.xlsx")

#################################
