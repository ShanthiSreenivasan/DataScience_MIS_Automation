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

LNT_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\LNT.xlsx") %>% select(Reference.ID,Journey.stage,Disbursal.date,Amount.chosen)

appos_map_LNT <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'LNT')
LNT_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('L&T Finance',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

#names(LNT_appos_dump)
#LNT_appos_dump<-replace(LNT_appos_dump$names, "L&T Finance", "LNT")
# %>% filter(grepl('L&T',names,ignore.case=TRUE)) 
########################LNT LenderFeedBack


colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "Journey.stage")] <- "customer_status_names"

colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "Reference.ID")] <- "mobile_number"


#LNT_lender_MIS$mobile_number <- as.numeric(LNT_lender_MIS$mobile_number)


LNT_lender_MIS$offer_application_number<-LNT_appos_dump$offer_application_number[match(LNT_lender_MIS$mobile_number, LNT_appos_dump$phone_home)]



LNT_lender_MIS$CURRENT_appops_status<-LNT_appos_dump$appops_status_code[match(LNT_lender_MIS$mobile_number, LNT_appos_dump$phone_home)]


LNT_lender_MIS$New_appops_status<-appos_map_LNT$New_Status[match(LNT_lender_MIS$customer_status_names, appos_map_LNT$CM_Status)]

LNT_lender_MIS$NEW_appops_description<-appos_map_LNT$Status_Description[match(LNT_lender_MIS$customer_status_names, appos_map_LNT$CM_Status)]



LNT_lender_MIS <-LNT_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
  ))

LNT_lender_MIS$Remark <- ifelse(is.na(LNT_lender_MIS$Remark), LNT_lender_MIS$New_appops_status, LNT_lender_MIS$Remark)


LNT_lender_MIS$Upload_Remarks <- paste(LNT_lender_MIS$customer_status_names,'-',LNT_lender_MIS$NEW_appops_description)

colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "offer_application_number")] <- "Application_Number"

LNT_upload<-LNT_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

LNT_upload<- LNT_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=LNT_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=LNT_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(LNT_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_LNT_FB_File.xlsx")

write.xlsx(LNT_upload, file = "C:\\R\\Lender Feedback\\Output\\LNT_upload.xlsx")

#################################
