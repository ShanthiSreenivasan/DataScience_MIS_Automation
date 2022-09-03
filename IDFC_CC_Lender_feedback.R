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

IDFC_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\IDFC_CC.xlsx", sheet = "Master file") %>% select(CRM.Lead.Id,Applicant.Mobile,Stage,Sub.Stage,Stage.by.Integration.Status,FI.Curing.Reason)


appos_map_IDFC <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'IDFC CC')
IDFC_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('IDFC',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

########################IDFC LenderFeedBack

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "Applicant.Mobile")] <- "mobile_number"

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "Sub.Stage")] <- "customer_status_name"

IDFC_lender_MIS$mobile_number <- as.numeric(IDFC_lender_MIS$mobile_number)


IDFC_lender_MIS$offer_application_number<-IDFC_appos_dump$offer_application_number[match(IDFC_lender_MIS$mobile_number, IDFC_appos_dump$phone_home)]



IDFC_lender_MIS$CURRENT_appops_status<-IDFC_appos_dump$appops_status_code[match(IDFC_lender_MIS$mobile_number, IDFC_appos_dump$phone_home)]


IDFC_lender_MIS$New_appops_status<-appos_map_IDFC$New_Status[match(IDFC_lender_MIS$customer_status_name, appos_map_IDFC$CM_Status)]

IDFC_lender_MIS$NEW_appops_description<-appos_map_IDFC$Status_Description[match(IDFC_lender_MIS$customer_status_name, appos_map_IDFC$CM_Status)]

names(IDFC_lender_MIS)
IDFC_lender_MIS <-IDFC_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_ CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not_in_ CRM',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

IDFC_lender_MIS$Remark <- ifelse(is.na(IDFC_lender_MIS$Remark), IDFC_lender_MIS$New_appops_status, IDFC_lender_MIS$Remark)

#IDFC_lender_MIS<-IDFC_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
IDFC_lender_MIS$Upload_Remarks <- paste(IDFC_lender_MIS$customer_status_name,'-',IDFC_lender_MIS$Stage.by.Integration.Status)

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "offer_application_number")] <- "Application_Number"

IDFC_upload<- IDFC_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not_in_ CRM")) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=IDFC_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=IDFC_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(IDFC_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_IDFC_FB_File.xlsx")

write.xlsx(IDFC_upload, file = "C:\\R\\Lender Feedback\\Output\\IDFC_upload.xlsx")

#################################
