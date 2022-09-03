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

SME_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\SME.xlsx", sheet='Sheet0') %>% select(Phone,Partner.Main.Status,Partner.Sub.Status,Approved.Amount,Reject.Remarks,TSE.Comments)




names(SME_lender_MIS)
appos_map_SME <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'SME')
SME_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Sme Corner',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

########################SME LenderFeedBack

SME_lender_MIS$Phone <- as.numeric(SME_lender_MIS$Phone)

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "Phone")] <- "mobile_number"

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "Partner.Main.Status")] <- "customer_status_name"

#SME_appos_dump$ph_no<-SME_lender_MIS$mobile_number[match(SME_appos_dump$phone_home, SME_lender_MIS$mobile_number)]

SME_lender_MIS$offer_application_number<-SME_appos_dump$offer_application_number[match(SME_lender_MIS$mobile_number, SME_appos_dump$phone_home)]



SME_lender_MIS$CURRENT_appops_status<-SME_appos_dump$appops_status_code[match(SME_lender_MIS$mobile_number, SME_appos_dump$phone_home)]

SME_lender_MIS$New_appops_status<-appos_map_SME$New_Status[match(SME_lender_MIS$customer_status_name, appos_map_SME$CM_Status)]

SME_lender_MIS$NEW_appops_description<-appos_map_SME$Status_Description[match(SME_lender_MIS$customer_status_name, appos_map_SME$CM_Status)]


SME_lender_MIS <-SME_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    is.na(offer_application_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

SME_lender_MIS$Remark <- ifelse(is.na(SME_lender_MIS$Remark), SME_lender_MIS$New_appops_status, SME_lender_MIS$Remark)

SME_lender_MIS$Upload_Remarks <- paste(SME_lender_MIS$Partner.Sub.Status,"-",SME_lender_MIS$Reject.Remarks,"-",SME_lender_MIS$TSE.Comments)


colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "offer_application_number")] <- "Application_Number"

SME_lender_MIS<-SME_lender_MIS %>% mutate('Lender'="Sme Corner")


write.xlsx(SME_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SME_FB_File.xlsx")

SME_upload_1<-SME_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))





SME_upload_1<- SME_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=SME_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=SME_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(SME_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SME_upload.xlsx")

