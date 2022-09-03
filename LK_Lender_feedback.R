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

LK_lender_MIS <- read.csv("C:\\R\\Lender Feedback\\Input\\LK.csv") %>% select(Lead.Id,Mobile.Number,Application.Id,Status,Substatus,Rejection.Reason)

#names(LK_lender_MIS)
appos_map_LK <- read_xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Lendingkart')
LK_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Lending',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################LK LenderFeedBack

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "Mobile.Number")] <- "mobile_number"

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "Substatus")] <- "customer_status_name"

LK_lender_MIS$mobile_number <- as.numeric(LK_lender_MIS$mobile_number)


LK_lender_MIS$offer_application_number<-LK_appos_dump$offer_application_number[match(LK_lender_MIS$mobile_number, LK_appos_dump$phone_home)]



LK_lender_MIS$CURRENT_appops_status<-LK_appos_dump$appops_status_code[match(LK_lender_MIS$mobile_number, LK_appos_dump$phone_home)]


LK_lender_MIS$New_appops_status<-appos_map_LK$New_Status[match(LK_lender_MIS$customer_status_name, appos_map_LK$CM_Status)]

LK_lender_MIS$NEW_appops_description<-appos_map_LK$Status_Description[match(LK_lender_MIS$customer_status_name, appos_map_LK$CM_Status)]


LK_lender_MIS <-LK_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

LK_lender_MIS$Remark <- ifelse(is.na(LK_lender_MIS$Remark), LK_lender_MIS$New_appops_status, LK_lender_MIS$Remark)

LK_lender_MIS$Upload_Remarks <- paste(LK_lender_MIS$Status,'-',LK_lender_MIS$customer_status_name)

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "offer_application_number")] <- "Application_Number"


LK_lender_MIS<-LK_lender_MIS %>% mutate('Lender'="Lendingkart")

write.xlsx(LK_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_LK_FB_File.xlsx")

LK_upload<-LK_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

LK_upload<- LK_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=LK_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=LK_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(LK_upload, file = "C:\\R\\Lender Feedback\\Output\\LK_upload.xlsx")

#################################
