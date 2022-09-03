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
#library(DBI)
#library(RPostgreSQL)
#library(RMySQL)
#library(logging)
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

if(!dir.exists(paste("./Input/",thisdate,sep="")))
{
  dir.create(paste("./Input/",thisdate,sep=""))
} 


DIS_lender_MIS <- read_xlsx("C:\\R\\Lender Feedback\\Input\\Lead_Disbursed_Data_Report.xlsx", sheet='file') %>% select(PRPSLNO,PERMOBILE,LNAMT)



PRO_lender_MIS <- read_xlsx("C:\\R\\Lender Feedback\\Input\\Lead_Data_Proposal_Report.xlsx", sheet='file') %>% select(MOBILENO,REMARKS,REASONFORREJECTION,USERREMARKS,LOANAMOUNT)


#DIS_lender_MIS
appos_map_SCUF <- read_xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'sriram1')
SCUF_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Shriram City Union',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

########################SCUF LenderFeedBack
#SCUF_appos_dump <- appos_dump %>% filter(name %in% c('Lending kart'))
#names(DIS_lender_MIS)

colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "PERMOBILE")] <- "mobile_number"


SCUF_appos_dump$ph_no<-DIS_lender_MIS$mobile_number[match(SCUF_appos_dump$phone_home, DIS_lender_MIS$mobile_number)]

DIS_lender_MIS$offer_application_number<-SCUF_appos_dump$offer_application_number[match(DIS_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]



DIS_lender_MIS$CURRENT_appops_status<-SCUF_appos_dump$appops_status_code[match(DIS_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]


DIS_lender_MIS$New_appops_status<- '990'

DIS_lender_MIS$NEW_appops_description<-"Loan Disbursed"


DIS_lender_MIS <-DIS_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_CRM',
    is.na(offer_application_number) ~ 'Not_in_CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

DIS_lender_MIS$Remark <- ifelse(is.na(DIS_lender_MIS$Remark), DIS_lender_MIS$New_appops_status, DIS_lender_MIS$Remark)


#DIS_lender_MIS$NEW_appops_description<-appos_map_SCUF$Description[match(DIS_lender_MIS$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]

DIS_lender_MIS$Upload_Remarks <- str_c(DIS_lender_MIS$PRPSLNO,' ',DIS_lender_MIS$NEW_appops_description,' ',DIS_lender_MIS$LNAMT)

#


colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "offer_application_number")] <- "Application_Number"


DIS_lender_MIS<-DIS_lender_MIS %>% mutate('Lender'="Shriram City Union")

write.xlsx(DIS_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SCUF_Disbursed_FB_File.xlsx")


# SCUF_upload_1<- DIS_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not_in_CRM")) %>% select(Application_Number) %>% 
#   mutate(`App_Ops_Status`=DIS_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=DIS_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(SCUF_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_1.xlsx")


scuf_df1<-DIS_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

SCUF_upload_1<- scuf_df1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=scuf_df1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=scuf_df1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=scuf_df1$LNAMT,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(SCUF_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_1.xlsx")

######################


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "MOBILENO")] <- "mobile_number"


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "REMARKS")] <- "customer_status_name"

PRO_lender_MIS$ph_no<-DIS_lender_MIS$mobile_number[match(PRO_lender_MIS$mobile_number, DIS_lender_MIS$mobile_number)]
PRO_lender_MIS$ph_no[is.na(PRO_lender_MIS$ph_no)] <- 'FALSE'

PRO_lender_MIS<-PRO_lender_MIS %>% filter(ph_no=='FALSE')

SCUF_appos_dump$ph_no<-PRO_lender_MIS$mobile_number[match(SCUF_appos_dump$phone_home, PRO_lender_MIS$mobile_number)]

PRO_lender_MIS$offer_application_number<-SCUF_appos_dump$offer_application_number[match(PRO_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]



PRO_lender_MIS$CURRENT_appops_status<-SCUF_appos_dump$appops_status_code[match(PRO_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]

PRO_lender_MIS$New_appops_status<-appos_map_SCUF$New_Status[match(PRO_lender_MIS$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]

PRO_lender_MIS$NEW_appops_description<-appos_map_SCUF$Status_Description[match(PRO_lender_MIS$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]

#PRO_lender_MIS$NEW_appops_description<-appos_map_SCUF$Status_Description[match(PRO_lender_MIS$New_appops_status, appos_map_SCUF$New_Status)]

names(PRO_lender_MIS)

PRO_lender_MIS <-PRO_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_CRM',
    is.na(offer_application_number) ~ 'Not_in_CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

PRO_lender_MIS$Remark <- ifelse(is.na(PRO_lender_MIS$Remark), PRO_lender_MIS$New_appops_status, PRO_lender_MIS$Remark)

PRO_lender_MIS$Remark[is.na(PRO_lender_MIS$Remark)] <- "Repeated_Feedback_cases"

PRO_lender_MIS$Upload_Remarks <- str_c(PRO_lender_MIS$customer_status_name,'-',PRO_lender_MIS$REASONFORREJECTION,'-',PRO_lender_MIS$USERREMARKS)

PRO_lender_MIS$Upload_Remarks <- ifelse(is.na(PRO_lender_MIS$Upload_Remarks), PRO_lender_MIS$customer_status_name, PRO_lender_MIS$Upload_Remarks)

colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "offer_application_number")] <- "Application_Number"

PRO_lender_MIS<-PRO_lender_MIS %>% mutate('Lender'="Shriram City Union")

write.xlsx(PRO_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SCUF_Proposal_FB_File.xlsx")

scuf_df2<-PRO_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

SCUF_upload_2<- scuf_df2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=scuf_df2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=scuf_df2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(SCUF_upload_2, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_2.xlsx")
