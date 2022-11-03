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
# library(DBI)
# library(RPostgreSQL)
# library(RMySQL)
# library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
#library(bit64)
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

Status_YES_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\YES_CC.xlsx", sheet = "Sheet1") %>% select(customer_contact,status,dedupe_status,policy_check_status,cibil_check_status,idv,ekyc_status,Dip_Status,DIP.REJECT.REASON)

#View(YES_lender_MIS)

appos_map_YES <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Yesback cc')

YES_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Yes Bank',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
YES_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Yes Bank',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)

YES_appos_dump<-rbind(YES_appos_dump1,YES_appos_dump2)
#names(YES_lender_MIS)
########################YES LenderFeedBack


dedupe_status_F<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('FAIL')) %>% select(customer_contact) %>% mutate(`status`="dedupe_status_fail")

policy_check_status_F<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('PASS','NA') & policy_check_status %in% c('FAIL')) %>% select(customer_contact) %>% mutate(`status`="policy_check_status_fail")

cibil_check_status_F<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('PASS','NA'), policy_check_status %in% c('PASS','NA') & cibil_check_status %in% c('FAIL')) %>% select(customer_contact) %>% mutate(`status`="cibil_check_status_fail")

idv_F<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('PASS','NA'), policy_check_status %in% c('PASS','NA'), cibil_check_status %in% c('PASS','NA') & idv %in% c('FAIL','Fail')) %>% select(customer_contact) %>% mutate(`status`="idv_fail")

#ekyc_status_F<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('PASS','NA'), policy_check_status %in% c('PASS','NA'), cibil_check_status %in% c('PASS','NA'), idv %in% c('PASS','NA') & ekyc_status %in% c('FAIL','Fail')) %>% select(customer_contact) %>% mutate(`status`="ekyc_status_fail")


#merge_df1<-rbind(dedupe_status_F,policy_check_status_F,ekyc_status_F,cibil_check_status_F,idv_F)#

merge_df1<-rbind(dedupe_status_F,policy_check_status_F,cibil_check_status_F,idv_F)#ekyc_status_F,

ekyc_status_P<-Status_YES_lender_MIS %>% filter(dedupe_status %in% c('PASS','NA'), policy_check_status %in% c('PASS','NA'), cibil_check_status %in% c('PASS','NA'), idv %in% c('PASS','NA') & ekyc_status %in% c('PASS','NA','[NULL]')) %>% select(customer_contact,status)

merge_df2<-rbind(ekyc_status_P,merge_df1)

YES_lender_MIS<-merge_df2

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "customer_contact")] <- "mobile_number"

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "status")] <- "customer_status_name"



YES_lender_MIS$mobile_number <- as.numeric(YES_lender_MIS$mobile_number)


YES_lender_MIS$offer_application_number<-YES_appos_dump$offer_application_number[match(YES_lender_MIS$mobile_number, YES_appos_dump$phone_home)]



YES_lender_MIS$CURRENT_appops_status<-YES_appos_dump$appops_status_code[match(YES_lender_MIS$mobile_number, YES_appos_dump$phone_home)]

YES_lender_MIS$New_appops_status<-appos_map_YES$New_Status[match(YES_lender_MIS$customer_status_name, appos_map_YES$CM_Status)]

YES_lender_MIS$NEW_appops_description<-appos_map_YES$Status_Description[match(YES_lender_MIS$customer_status_name, appos_map_YES$CM_Status)]

#########################
YES_lender_MIS <-YES_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,690,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

YES_lender_MIS$Remark <- ifelse(is.na(YES_lender_MIS$Remark), YES_lender_MIS$New_appops_status, YES_lender_MIS$Remark)

#YES_lender_MIS<-YES_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
YES_lender_MIS$Upload_Remarks <- paste(YES_lender_MIS$NEW_appops_description,'-',YES_lender_MIS$customer_status_name)


YES_lender_MIS$Remark[YES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',YES_lender_MIS$New_appops_status,ignore.case=TRUE) & YES_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, YES_lender_MIS$New_appops_status %in% c(280,380) 

YES_lender_MIS$Remark[YES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',YES_lender_MIS$New_appops_status,ignore.case=TRUE) & YES_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, YES_lender_MIS$New_appops_status %in% c(480) 

YES_lender_MIS$Remark[YES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',YES_lender_MIS$New_appops_status,ignore.case=TRUE) & YES_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, YES_lender_MIS$New_appops_status %in% c(580) 

YES_lender_MIS$Remark[YES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',YES_lender_MIS$New_appops_status,ignore.case=TRUE) & YES_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, YES_lender_MIS$New_appops_status %in% c(680) 

colnames(YES_lender_MIS)[which(names(YES_lender_MIS) == "offer_application_number")] <- "Application_Number"

YES_lender_MIS<-YES_lender_MIS %>% mutate('Lender'="Yes Bank",`loan_amount`='',`Customer_requested_loan_amount`='')


YES_upload<- YES_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not in CRM",'Stuck_cases')) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=YES_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=YES_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(YES_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_YES_FB_File.xlsx")

#write.xlsx(YES_upload, file = "C:\\R\\Lender Feedback\\Output\\YES_upload.xlsx")

#write.xlsx(YES_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\YES_appos_dump.xlsx")

#################################
