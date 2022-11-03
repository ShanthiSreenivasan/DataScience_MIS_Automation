rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)
library(xtable)
library(data.table)

library(stringr)
library(lubridate)
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

appos_map_HC <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx", sheet = 'HomeCredit')

HC_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('HOME CREDIT',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
HC_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('HOME CREDIT',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

HC_appos_dump<-rbind(HC_appos_dump1,HC_appos_dump2)
######################



Status_HC_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\HC.xlsx") %>% select(MOBILE,DECISION,LOAN_AMOUNT,APPLICATION_DATE,LOAN_PURPOSE_DATE,
                                                                                       PROFILE_DATE,FIRST_SUBMISSION_DATE,PRODUCT_SELECTION_DATE,BANK_CONSENT_DATE,
                                                                                       BANKVERIFICATION_DATE,FINAL_SUBMISSION_DATE,APPROVED_DATE,SIGNATURE_DATE,PAID_DATE)


colnames(Status_HC_lender_MIS)[which(names(Status_HC_lender_MIS) == "LOAN_AMOUNT")] <- "loan_amount"

APPLI<-Status_HC_lender_MIS %>% filter(!APPLICATION_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="APPLICATION_DATE")


LOAN<-Status_HC_lender_MIS %>% filter(!LOAN_PURPOSE_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="LOAN_PURPOSE_DATE")

PROFILE<-Status_HC_lender_MIS %>% filter(!PROFILE_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="PROFILE_DATE")

f_Sub<-Status_HC_lender_MIS %>% filter(!FIRST_SUBMISSION_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="FIRST_SUBMISSION_DATE")

PRODUCT<-Status_HC_lender_MIS %>% filter(!PRODUCT_SELECTION_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="PRODUCT_SELECTION_DATE")

#BANK<-Status_HC_lender_MIS %>% filter(!BANK_CONSENT_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="BANK_CONSENT_DATE")

#BANK_VER<-Status_HC_lender_MIS %>% filter(!BANKVERIFICATION_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="FINAL_SUBMISSION_DATE")

FINA<-Status_HC_lender_MIS %>% filter(!FINAL_SUBMISSION_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="BANK_CONSENT_DATE")

APP<-Status_HC_lender_MIS %>% filter(!APPROVED_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="APPROVED_DATE")

SIG<-Status_HC_lender_MIS %>% filter(!SIGNATURE_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="SIGNATURE_DATE")

PAID<-Status_HC_lender_MIS %>% filter(!PAID_DATE == 'NA') %>% select(MOBILE,loan_amount) %>% mutate(`REMARKS`="PAID_DATE")

HC_lender_MIS<-rbind(APPLI,LOAN,PROFILE,f_Sub,PRODUCT,FINA,APP,SIG,PAID)#BANK,BANK_VER,


colnames(HC_lender_MIS)[which(names(HC_lender_MIS) == "MOBILE")] <- "mobile_number"


colnames(HC_lender_MIS)[which(names(HC_lender_MIS) == "REMARKS")] <- "customer_status_name"


#HC_lender_MIS$mobile_number <- as.numeric(HC_lender_MIS$mobile_number)



HC_lender_MIS$offer_application_number<-HC_appos_dump$offer_application_number[match(HC_lender_MIS$mobile_number, HC_appos_dump$phone_home)]



HC_lender_MIS$CURRENT_appops_status<-HC_appos_dump$appops_status_code[match(HC_lender_MIS$mobile_number, HC_appos_dump$phone_home)]

HC_lender_MIS$New_appops_status<-appos_map_HC$New_Status[match(HC_lender_MIS$customer_status_name, appos_map_HC$CM_Status)]

HC_lender_MIS$NEW_appops_description<-appos_map_HC$Status_Description[match(HC_lender_MIS$customer_status_name, appos_map_HC$CM_Status)]


HC_lender_MIS <-HC_lender_MIS %>%
  mutate(Remark = case_when(
    #CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

HC_lender_MIS$Remark <- ifelse(is.na(HC_lender_MIS$Remark), HC_lender_MIS$New_appops_status, HC_lender_MIS$Remark)



HC_lender_MIS$Upload_Remarks <- str_c(HC_lender_MIS$customer_status_name,'-',HC_lender_MIS$NEW_appops_description)


HC_lender_MIS$Remark[HC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',HC_lender_MIS$New_appops_status,ignore.case=TRUE) & HC_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, HC_lender_MIS$New_appops_status %in% c(280,380) 

HC_lender_MIS$Remark[HC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',HC_lender_MIS$New_appops_status,ignore.case=TRUE) & HC_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, HC_lender_MIS$New_appops_status %in% c(480) 

HC_lender_MIS$Remark[HC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',HC_lender_MIS$New_appops_status,ignore.case=TRUE) & HC_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, HC_lender_MIS$New_appops_status %in% c(580) 

HC_lender_MIS$Remark[HC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',HC_lender_MIS$New_appops_status,ignore.case=TRUE) & HC_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, HC_lender_MIS$New_appops_status %in% c(680) 

colnames(HC_lender_MIS)[which(names(HC_lender_MIS) == "offer_application_number")] <- "Application_Number"

HC_lender_MIS<-HC_lender_MIS %>% mutate('Lender'="HomeCredit",`Customer_requested_loan_amount`='')


write.xlsx(HC_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_HC_FB_File.xlsx")

HC_upload<-HC_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))


HC_upload<- HC_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=HC_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=HC_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(HC_upload, file = "C:\\R\\Lender Feedback\\Output\\HC_upload.xlsx")

#################################
