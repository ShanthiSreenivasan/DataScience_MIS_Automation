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

ES_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\ES.xlsx") %>% select(mobile_number,status,subs_status,app_download_flag,rejectreasonsgroup,app_download_flag,first_disb_loan_amt)

#names(ES_lender_MIS)
appos_map_ES <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'EarlySalary')

ES_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Fibe',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
ES_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Fibe',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
ES_appos_dump<-rbind(ES_appos_dump1,ES_appos_dump2)

########################ES LenderFeedBack


#colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "subs_status")] <- "customer_status_name"

ES_lender_MIS$mobile_number <- as.numeric(ES_lender_MIS$mobile_number)

ES_lender_MIS$app_download_flag <- as.numeric(ES_lender_MIS$app_download_flag)

appos_map_ES$app_download_flag <- as.numeric(appos_map_ES$app_download_flag)

ES_lender_MIS$offer_application_number<-ES_appos_dump$offer_application_number[match(ES_lender_MIS$mobile_number, ES_appos_dump$phone_home)]



ES_lender_MIS$CURRENT_appops_status<-ES_appos_dump$appops_status_code[match(ES_lender_MIS$mobile_number, ES_appos_dump$phone_home)]


ES_lender_MIS<-left_join(ES_lender_MIS,appos_map_ES, by = c('subs_status','app_download_flag'))

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "New_Status")] <- "New_appops_status"

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "Status_Description")] <- "NEW_appops_description"

#ES_lender_MIS$New_appops_status<-appos_map_ES$New_Status[match(ES_lender_MIS$subs_status, appos_map_ES$subs_status)]

#ES_lender_MIS$NEW_appops_description<-appos_map_ES$Status_Description[match(ES_lender_MIS$subs_status, appos_map_ES$subs_status)]

#names(ES_lender_MIS)


ES_lender_MIS <-ES_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

ES_lender_MIS$Remark <- ifelse(is.na(ES_lender_MIS$Remark), ES_lender_MIS$New_appops_status, ES_lender_MIS$Remark)




ES_lender_MIS$Upload_Remarks <- paste(ES_lender_MIS$status,'-',ES_lender_MIS$rejectreasonsgroup,'-',ES_lender_MIS$NEW_appops_description)

ES_lender_MIS$Remark[ES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',ES_lender_MIS$New_appops_status,ignore.case=TRUE) & ES_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, ES_lender_MIS$New_appops_status %in% c(280,380) 

ES_lender_MIS$Remark[ES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',ES_lender_MIS$New_appops_status,ignore.case=TRUE) & ES_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, ES_lender_MIS$New_appops_status %in% c(480) 

ES_lender_MIS$Remark[ES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',ES_lender_MIS$New_appops_status,ignore.case=TRUE) & ES_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, ES_lender_MIS$New_appops_status %in% c(580) 

ES_lender_MIS$Remark[ES_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',ES_lender_MIS$New_appops_status,ignore.case=TRUE) & ES_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, ES_lender_MIS$New_appops_status %in% c(680) 

ES_lender_MIS$Remark[grepl('Journey Reject - Salary &lt;18000',ES_lender_MIS$customer_status_name,ignore.case=TRUE) & ES_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, ES_lender_MIS$New_appops_status %in% c(680) 

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "subs_status")] <- "customer_status_name"

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(ES_lender_MIS)[which(names(ES_lender_MIS) == "first_disb_loan_amt")] <- "loan_amount"

ES_lender_MIS<-ES_lender_MIS %>% mutate('Lender'="Early Salary",`Customer_requested_loan_amount`='')


write.xlsx(ES_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_ES_FB_File.xlsx")

ES_upload<-ES_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

ES_upload<- ES_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=ES_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=ES_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(ES_upload, file = "C:\\R\\Lender Feedback\\Output\\ES_upload.xlsx")

#################################
