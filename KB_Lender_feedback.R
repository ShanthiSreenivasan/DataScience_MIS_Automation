rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)
library(xtable)
library(stringr)
library(lubridate)
library(data.table)
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


KB_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\KB.xlsx", sheet = "CM_Base") %>% select(uId,State,first_loan_taken_date,first_loan_gmv,mobile,user_subState,Rejection_Reason)


#DIS_KB_lender_MIS
appos_map_KB <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'KREDITBEE')
KB_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Kredit Bee',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

########################KB LenderFeedBack

KB_lender_MIS$mobile <- as.numeric(KB_lender_MIS$mobile)

#names(KB_lender_MIS)

colnames(KB_lender_MIS)[which(names(KB_lender_MIS) == "mobile")] <- "mobile_number"


colnames(KB_lender_MIS)[which(names(KB_lender_MIS) == "user_subState")] <- "CM_Status"


KB_lender_MIS$offer_application_number<-KB_appos_dump$offer_application_number[match(KB_lender_MIS$mobile_number, KB_appos_dump$phone_home)]



KB_lender_MIS$CURRENT_appops_status<-KB_appos_dump$appops_status_code[match(KB_lender_MIS$mobile_number, KB_appos_dump$phone_home)]


KB_df<-left_join(KB_lender_MIS,appos_map_KB, by = c('State','CM_Status'))

colnames(KB_df)[which(names(KB_df) == "New_Status")] <- "New_appops_status"

#KB_lender_MIS$New_appops_status<-appos_map_KB$New_Status[match(KB_lender_MIS$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS$New_appops_status, appos_map_KB$New_Status)]
#KB_df$latest_month <- format(as.Date(KB_df$first_loan_taken_date, format="%d/%m/%Y"),"%m")

#View(KB_lender_MIS)

current_mon<-format(as.Date(Sys.Date(), format="%d/%m/%Y"),"%m")

#subset(KB_df, first_loan_taken_date > Sys.Date() - months(2))

KB_df<-KB_df %<>%
  mutate(first_loan_taken_date_trim= as.Date(first_loan_taken_date, format= "%d/%m/%Y"))


#KB_df$first_loan_taken_date <- as.Date(KB_df$first_loan_taken_date, format = "%d/%m/%Y")

#class(KB_df$first_loan_taken_date)

KB_df <-KB_df %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_CRM',
    is.na(offer_application_number) ~ 'Not_in_CRM',
    grepl('Confirmed', State, ignore.case = T) & as.numeric(latest_month == current_mon) ~ '990',# & grepl('Yes', first_loan_taken_or_not, ignore.case = T) ~ '990',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    #(New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

KB_df$Remark <- ifelse(is.na(KB_df$Remark), KB_df$New_appops_status, KB_df$Remark)

KB_df$Upload_Remarks <- str_c(KB_df$State,'-',KB_df$CM_Status,'-',KB_df$Rejection_Reason,'-',KB_df$NEW_appops_description)

#KB_df<-KB_df %>% filter(!is.na(offer_application_number))

#############to chk
#KB_df<-KB_df %>% filter(!is.na(offer_application_number) & !is.na(Remark))




colnames(KB_df)[which(names(KB_df) == "offer_application_number")] <- "Application_Number"

KB_df<-KB_df %>% mutate('Lender'="Kredit Bee")


write.xlsx(KB_df, file = "C:\\R\\Lender Feedback\\Output\\Revised_KB_FB_File.xlsx")

KB_upload_1<-KB_df %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

KB_upload_1<- KB_upload_1 %>% filter(!grepl('Repeated_Feedback_cases, Not_in_CRM',Remark,ignore.case=TRUE)) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=KB_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=KB_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=KB_upload_1$first_loan_gmv,`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



write.xlsx(KB_upload_1, file = "C:\\R\\Lender Feedback\\Output\\KB_upload_1.xlsx")

