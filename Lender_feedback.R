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

LastMonth_API <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'LastMonth_API') %>% select(customer_id,mobile_no,customer_status_name,loan_amount)

CurMonth_API <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'CurMonth_API') %>% select(customer_id,mobile_no,customer_status_name,loan_amount)


LastMonth_Notregister <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'LastMonth_Notregister') %>% select(mobile)


CurMonth_Notregister <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'CurMonth_Notregister') %>% select(mobile)

appos_map_CashE <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Cashe')
appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(name %in% c('CashE')) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################CashE LenderFeedBack
CashE_appos_dump <- appos_dump %>% filter(name %in% c('CashE'))

CashE_appos_dump$ph_no<-LastMonth_API$mobile_no[match(CashE_appos_dump$phone_home, LastMonth_API$mobile_no)]

#names(LastMonth_API)

LastMonth_API$offer_application_number<-CashE_appos_dump$offer_application_number[match(LastMonth_API$mobile_no, CashE_appos_dump$phone_home)]



LastMonth_API$CURRENT_appops_status<-CashE_appos_dump$appops_status_code[match(LastMonth_API$mobile_no, CashE_appos_dump$phone_home)]


LastMonth_API$New_appops_status<-appos_map_CashE$New_Status[match(LastMonth_API$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]

#LastMonth_API$NEW_appops_description<-appos_map_CashE$New_Status[match(LastMonth_API$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]

LastMonth_API$NEW_appops_description<-appos_map_CashE$Status_Description[match(LastMonth_API$New_appops_status, appos_map_CashE$New_Status)]



Not_in_CRM_cases <- CashE_appos_dump %>% filter(is.na(phone_home))


LastMonth_API_fbNOTrcvd_cases<-LastMonth_API %>% filter(is.na(offer_application_number))

LastMonth_API<-LastMonth_API %>% filter(!is.na(offer_application_number))
#feedback_file = appos_dump %>% left_join(appos_dump$offer_application_number, by = c('phone_home'))

LastMonth_API<-LastMonth_API %>%
  mutate(Repeated_Feedback_cases = case_when(
    CURRENT_appops_status == New_appops_status ~ 1))

LastMonth_API$Repeated_Feedback_cases[is.na(LastMonth_API$Repeated_Feedback_cases)] <- 0

LastMonth_API<-LastMonth_API %>% filter(!is.na(offer_application_number), (New_appops_status>CURRENT_appops_status), Repeated_Feedback_cases==0)

    
#IVR_tamil$ph_no<-TCN_dump$Phone.Number[match(IVR_tamil$phone_home, TCN_dump$Phone.Number)]

#LastMonth_API_new<-LastMonth_API %>% filter(New_appops_status>CURRENT_appops_status & Repeated_Feedback_cases==0)

LastMonth_API_new_rej <-LastMonth_API %>%
  mutate(New_status = case_when(
    !is.na(loan_amount) ~ '990',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580',
    (New_appops_status ==990) ~ '990'
    ))

#LastMonth_API_new_rej$New_status[is.na(LastMonth_API_new_rej$New_status)] <- `LastMonth_API$New_appops_status`

write.xlsx(LastMonth_API_new, file = "C:\\R\\Lender Feedback\\CashE_upload-1.xlsx")

write.xlsx(LastMonth_API_new_rej, file = "C:\\R\\Lender Feedback\\CashE_upload-2.xlsx")

names(LastMonth_API)

desc_LastMonth_API <- LastMonth_API_new_rej %>% group_by(New_appops_status) %>% dplyr::summarise(Total = n_distinct(mobile_no, na.rm = TRUE)) %>% adorn_totals("row")


CashE_upload<-read.xlsx("C:\\R\\Lender Feedback\\CashE_upload-1.xlsx")

CashE_upload<- LastMonth_API_new_rej %>% filter(NEW_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,NEW_appops_description) %>% 
  mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

list_of_datasets <- list("FeedBack_update" = LastMonth_API_new_rej, "FeedBack_summary" = desc_LastMonth_API)
write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\CashE_feedback.xlsx")

write.xlsx(CashE_upload, file = "C:\\R\\Lender Feedback\\CashE_upload.xlsx")


###########################
CashE_appos_dump_2 <- appos_dump %>% filter(name %in% c('CashE'))

CashE_appos_dump_2$ph_no<-CurMonth_API$mobile_no[match(CashE_appos_dump_2$phone_home, CurMonth_API$mobile_no)]

#names(CurMonth_API)

CurMonth_API$offer_application_number<-CashE_appos_dump$offer_application_number[match(CurMonth_API$mobile_no, CashE_appos_dump_2$phone_home)]



CurMonth_API$CURRENT_appops_status<-CashE_appos_dump$appops_status_code[match(CurMonth_API$mobile_no, CashE_appos_dump_2$phone_home)]


CurMonth_API$New_appops_status<-appos_map_CashE$New_Status[match(CurMonth_API$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]

#CurMonth_API$NEW_appops_description<-appos_map_CashE$New_Status[match(CurMonth_API$customer_status_name, appos_map_CashE$Campaign_Sub_Status)]

CurMonth_API$NEW_appops_description<-appos_map_CashE$Status_Description[match(CurMonth_API$New_appops_status, appos_map_CashE$New_Status)]



Not_in_CRM_cases <- CashE_appos_dump %>% filter(is.na(phone_home))


CurMonth_API_fbNOTrcvd_cases<-CurMonth_API %>% filter(is.na(offer_application_number))

CurMonth_API<-CurMonth_API %>% filter(!is.na(offer_application_number))
#feedback_file = appos_dump %>% left_join(appos_dump$offer_application_number, by = c('phone_home'))

CurMonth_API<-CurMonth_API %>%
  mutate(Repeated_Feedback_cases = case_when(
    CURRENT_appops_status == New_appops_status ~ 1))

CurMonth_API$Repeated_Feedback_cases[is.na(CurMonth_API$Repeated_Feedback_cases)] <- 0

CurMonth_API<-CurMonth_API %>% filter(!is.na(offer_application_number), (New_appops_status>CURRENT_appops_status), Repeated_Feedback_cases==0)


#IVR_tamil$ph_no<-TCN_dump$Phone.Number[match(IVR_tamil$phone_home, TCN_dump$Phone.Number)]

#CurMonth_API_new<-CurMonth_API %>% filter(New_appops_status>CURRENT_appops_status & Repeated_Feedback_cases==0)

CurMonth_API_new_rej <-CurMonth_API %>%
  mutate(New_status = case_when(
    !is.na(loan_amount) ~ '990',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580',
    (New_appops_status ==990) ~ '990'
  ))

#CurMonth_API_new_rej$New_status[is.na(CurMonth_API_new_rej$New_status)] <- `CurMonth_API$New_appops_status`

write.xlsx(CurMonth_API_new_rej, file = "C:\\R\\Lender Feedback\\CashE_upload_2.xlsx")

names(CurMonth_API)

desc_CurMonth_API <- CurMonth_API_new_rej %>% group_by(New_appops_status) %>% dplyr::summarise(Total = n_distinct(mobile_no, na.rm = TRUE)) %>% adorn_totals("row")


CashE_upload_2<-read.xlsx("C:\\R\\Lender Feedback\\CashE_upload-2.xlsx")

CashE_upload_2<- CurMonth_API_new_rej %>% filter(NEW_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,NEW_appops_description) %>% 
  mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

list_of_datasets <- list("FeedBack_update" = CurMonth_API_new_rej, "FeedBack_summary" = desc_CurMonth_API)
write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\CashE_feedback-2.xlsx")

write.xlsx(CashE_upload_2, file = "C:\\R\\Lender Feedback\\CashE_upload_2.xlsx")


write.xlsx(LastMonth_Notregister, file = "C:\\R\\Lender Feedback\\CashE_upload_3.xlsx")

write.xlsx(CurMonth_Notregister, file = "C:\\R\\Lender Feedback\\CashE_upload_4.xlsx")


########################KB LenderFeedBack

KB_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\KB.xlsx",sheet = 'Sheet1') %>% select(uId,State,user_subState,latest_loan_state,first_loan_gmv,mobile)
appos_map_KB <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Krazy Bee')

KB_appos_dump <- appos_dump %>% filter(name %in% c('Kredit Bee')) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

KB_lender_MIS$offer_application_number<-KB_appos_dump$offer_application_number[match(KB_lender_MIS$mobile, KB_appos_dump$phone_home)]

KB_lender_MIS$appops_status<-KB_appos_dump$appops_status_code[match(KB_lender_MIS$mobile, KB_appos_dump$phone_home)]

KB_lender_MIS <-KB_lender_MIS %>%
  mutate(New_appops_status = case_when(
    !is.na(first_loan_gmv) ~ "990",
    appops_status >280 & appops_status <=380  ~ "480",
    appops_status >380 & appops_status <=500  ~ "490",
    appops_status >500 & appops_status <=590  ~ "590",
    appops_status ==990 ~ "990"
  ))


KB_lender_MIS$New_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS$New_appops_status, appos_map_KB$New.Status)]


KB_desc_lender_MIS <- KB_lender_MIS %>% group_by(New_appops_description) %>% dplyr::summarise(Total = n_distinct(uId, na.rm = TRUE)) %>% adorn_totals("row")


KB_upload<- KB_lender_MIS %>% filter(New_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,New_appops_description) %>% 
  mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

list_of_datasets <- list("FeedBack_update" = KB_lender_MIS, "FeedBack_summary" = KB_desc_lender_MIS)
write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\KB_feedback.xlsx")

write.xlsx(KB_upload, file = "C:\\R\\Lender Feedback\\KB_upload.xlsx")

