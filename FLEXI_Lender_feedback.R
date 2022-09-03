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

approved<- read.xlsx("C:\\R\\Lender Feedback\\Input\\FLEXI.xlsx", sheet='Approved') %>% select(mobile_no,s1,s2,s3,Loan.Amount)

inprogress<- read.xlsx("C:\\R\\Lender Feedback\\Input\\FLEXI.xlsx", sheet='Inprogress') %>% select(mobile_no,s1,s2,s3,Loan.Amount)

rejected<- read.xlsx("C:\\R\\Lender Feedback\\Input\\FLEXI.xlsx", sheet='Rejected') %>% select(mobile_no,s1,s2,s3,Loan.Amount)

FLEXI_lender_MIS<-rbind(approved,inprogress,rejected)


names(FLEXI_lender_MIS)
appos_map_FLEXI <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Flexiloan')
FLEXI_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('FLEXI',name,ignore.case=TRUE)) %>% select(leadid,phone_home,offer_application_number,status,name,appops_status_code)

########################FLEXI LenderFeedBack

#FLEXI_lender_MIS$mobile_no <- as.numeric(FLEXI_lender_MIS$mobile_no)

colnames(FLEXI_lender_MIS)[which(names(FLEXI_lender_MIS) == "mobile_no")] <- "mobile_number"

colnames(FLEXI_lender_MIS)[which(names(FLEXI_lender_MIS) == "s1")] <- "State"

colnames(FLEXI_lender_MIS)[which(names(FLEXI_lender_MIS) == "s2")] <- "CM_Status"


#FLEXI_appos_dump$ph_no<-FLEXI_lender_MIS$mobile_number[match(FLEXI_appos_dump$phone_home, FLEXI_lender_MIS$mobile_number)]

FLEXI_lender_MIS$offer_application_number<-FLEXI_appos_dump$offer_application_number[match(FLEXI_lender_MIS$mobile_number, FLEXI_appos_dump$phone_home)]



FLEXI_lender_MIS$CURRENT_appops_status<-FLEXI_appos_dump$appops_status_code[match(FLEXI_lender_MIS$mobile_number, FLEXI_appos_dump$phone_home)]


fle_df<-left_join(FLEXI_lender_MIS,appos_map_FLEXI, by = c('State','CM_Status'))

colnames(fle_df)[which(names(fle_df) == "New_Status")] <- "New_appops_status"


fle_df$New_appops_status<-appos_map_FLEXI$New_Status[match(fle_df$CM_Status, appos_map_FLEXI$CM_Status)]

fle_df$NEW_appops_description<-appos_map_FLEXI$Status_Description[match(fle_df$CM_Status, appos_map_FLEXI$CM_Status)]


fle_df <-fle_df %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    is.na(offer_application_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

fle_df$Remark <- ifelse(is.na(fle_df$Remark), fle_df$New_appops_status, fle_df$Remark)

fle_df$Upload_Remarks <- paste(fle_df$Partner.Sub.Status,"-",fle_df$Reject.Remarks,"-",fle_df$TSE.Comments)


colnames(fle_df)[which(names(fle_df) == "offer_application_number")] <- "Application_Number"

fle_df<-fle_df %>% mutate('Lender'="FLEXI Corner")


write.xlsx(fle_df, file = "C:\\R\\Lender Feedback\\Output\\Revised_FLEXI_FB_File.xlsx")

FLEXI_upload_1<-fle_df %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))





FLEXI_upload_1<- FLEXI_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=FLEXI_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=FLEXI_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(FLEXI_upload_1, file = "C:\\R\\Lender Feedback\\Output\\FLEXI_upload.xlsx")

