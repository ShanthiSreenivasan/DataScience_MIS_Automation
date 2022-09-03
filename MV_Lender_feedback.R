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

DIS_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV1.xlsx", sheet = "disbursals") %>% select(mobile_number,loan_amount)


#PRO_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV.xlsx", sheet = "Sheet1")

#DIS_MV_lender_MIS
appos_map_MV <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx", sheet = 'MoneyView')
MV_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Money View',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

########################MV LenderFeedBack

DIS_MV_lender_MIS$loan_amount <- as.numeric(DIS_MV_lender_MIS$loan_amount)

#colnames(DIS_MV_lender_MIS)[which(names(DIS_MV_lender_MIS) == "PERMOBILE")] <- "mobile_number"


MV_appos_dump$ph_no<-DIS_MV_lender_MIS$mobile_number[match(MV_appos_dump$phone_home, DIS_MV_lender_MIS$mobile_number)]

DIS_MV_lender_MIS$offer_application_number<-MV_appos_dump$offer_application_number[match(DIS_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]



DIS_MV_lender_MIS$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(DIS_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]


DIS_MV_lender_MIS$NEW_appops_status<- '990'

DIS_MV_lender_MIS$NEW_appops_description<-"Loan Disbursed"

#DIS_MV_lender_MIS$NEW_appops_description<-appos_map_MV$Description[match(DIS_MV_lender_MIS$customer_status_name, appos_map_MV$Campaign_Sub_Status)]

DIS_MV_lender_MIS <-DIS_MV_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    NEW_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (NEW_appops_status ==990) ~ '990',
    (NEW_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (NEW_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (NEW_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

DIS_MV_lender_MIS$Remark <- ifelse(is.na(DIS_MV_lender_MIS$Remark), DIS_MV_lender_MIS$NEW_appops_status, DIS_MV_lender_MIS$Remark)


DIS_MV_lender_MIS$Upload_Remarks <- str_c(DIS_MV_lender_MIS$NEW_appops_description,' ',DIS_MV_lender_MIS$loan_amount)


colnames(DIS_MV_lender_MIS)[which(names(DIS_MV_lender_MIS) == "offer_application_number")] <- "Application_Number"

DIS_MV_lender_MIS<-DIS_MV_lender_MIS %>% mutate('Lender'="Money View")

write.xlsx(DIS_MV_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Disbursed_FB_File_1.xlsx")

MV_upload_1<-DIS_MV_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))


MV_upload_1<- MV_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=MV_upload_1$loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



write.xlsx(MV_upload_1, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_1_1.xlsx")


######################



Status_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV1.xlsx", sheet = "status") %>% select(Mobile,Offer_selected,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved)

offer<-Status_MV_lender_MIS %>% filter(Offer_selected==1) %>% select(Mobile) %>% mutate(`REMARKS`="Offer_selected")


Rejected<-Status_MV_lender_MIS %>% filter(Rejected==1) %>% select(Mobile) %>% mutate(`REMARKS`="Rejected")

Cancelled<-Status_MV_lender_MIS %>% filter(Cancelled==1) %>% select(Mobile) %>% mutate(`REMARKS`="Cancelled")

Submits<-Status_MV_lender_MIS %>% filter(Submits==1) %>% select(Mobile) %>% mutate(`REMARKS`="Submits")

In_Process<-Status_MV_lender_MIS %>% filter(In_Process==1) %>% select(Mobile) %>% mutate(`REMARKS`="In_Process")

NACH<-Status_MV_lender_MIS %>% filter(NACH==1) %>% select(Mobile) %>% mutate(`REMARKS`="NACH")

Loan_Approved<-Status_MV_lender_MIS %>% filter(Loan_Approved==1) %>% select(Mobile) %>% mutate(`REMARKS`="Loan_Approved")


PRO_MV_lender_MIS<-rbind(offer,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved)

colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "Mobile")] <- "mobile_number"


colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "REMARKS")] <- "customer_status_name"


#PRO_MV_lender_MIS$mobile_number <- as.numeric(PRO_MV_lender_MIS$mobile_number)


PRO_MV_lender_MIS$ph_no<-DIS_MV_lender_MIS$mobile_number[match(PRO_MV_lender_MIS$mobile_number, DIS_MV_lender_MIS$mobile_number)]
PRO_MV_lender_MIS$ph_no[is.na(PRO_MV_lender_MIS$ph_no)] <- 'FALSE'

PRO_MV_lender_MIS<-PRO_MV_lender_MIS %>% filter(ph_no=='FALSE')

MV_appos_dump$ph_no<-PRO_MV_lender_MIS$mobile_number[match(MV_appos_dump$phone_home, PRO_MV_lender_MIS$mobile_number)]

PRO_MV_lender_MIS$offer_application_number<-MV_appos_dump$offer_application_number[match(PRO_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]



PRO_MV_lender_MIS$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(PRO_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]

PRO_MV_lender_MIS$NEW_appops_status<-appos_map_MV$New_Status[match(PRO_MV_lender_MIS$customer_status_name, appos_map_MV$CM_Status)]

PRO_MV_lender_MIS$NEW_appops_description<-appos_map_MV$Status_Description[match(PRO_MV_lender_MIS$customer_status_name, appos_map_MV$CM_Status)]


PRO_MV_lender_MIS <-PRO_MV_lender_MIS %>%
  mutate(Remark = case_when(
    is.na(offer_application_number) ~ 'Not in CRM',
    NEW_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (NEW_appops_status ==990) ~ '990',
    (NEW_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (NEW_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (NEW_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

PRO_MV_lender_MIS$Remark <- ifelse(is.na(PRO_MV_lender_MIS$Remark), PRO_MV_lender_MIS$NEW_appops_status, PRO_MV_lender_MIS$Remark)



PRO_MV_lender_MIS$Upload_Remarks <- str_c(PRO_MV_lender_MIS$customer_status_name,'-',PRO_MV_lender_MIS$NEW_appops_description)

colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "offer_application_number")] <- "Application_Number"


PRO_MV_lender_MIS<-PRO_MV_lender_MIS %>% mutate('Lender'="Money View")

write.xlsx(PRO_MV_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Proposal_FB_File_1.xlsx") 

MV_upload_2<-PRO_MV_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

MV_upload_2<- MV_upload_2 %>% filter(!grepl('Repeated_Feedback_cases, Not_in_CRM',Remark,ignore.case=TRUE)) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(MV_upload_2, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_2_1.xlsx")

##########################################################################



DIS_MV_lender_MIS_2 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV2.xlsx", sheet = "disbursals") %>% select(mobile_number,loan_amount)

########################MV LenderFeedBack

DIS_MV_lender_MIS_2$loan_amount <- as.numeric(DIS_MV_lender_MIS_2$loan_amount)

#colnames(DIS_MV_lender_MIS_2)[which(names(DIS_MV_lender_MIS_2) == "PERMOBILE")] <- "mobile_number"


MV_appos_dump$ph_no<-DIS_MV_lender_MIS_2$mobile_number[match(MV_appos_dump$phone_home, DIS_MV_lender_MIS_2$mobile_number)]

DIS_MV_lender_MIS_2$offer_application_number<-MV_appos_dump$offer_application_number[match(DIS_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]



DIS_MV_lender_MIS_2$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(DIS_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]


DIS_MV_lender_MIS_2$NEW_appops_status<- '990'

DIS_MV_lender_MIS_2$NEW_appops_description<-"Loan Disbursed"

#DIS_MV_lender_MIS_2$NEW_appops_description<-appos_map_MV$Description[match(DIS_MV_lender_MIS_2$customer_status_name, appos_map_MV$Campaign_Sub_Status)]

DIS_MV_lender_MIS_2 <-DIS_MV_lender_MIS_2 %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    NEW_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (NEW_appops_status ==990) ~ '990',
    (NEW_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (NEW_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (NEW_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

DIS_MV_lender_MIS_2$Remark <- ifelse(is.na(DIS_MV_lender_MIS_2$Remark), DIS_MV_lender_MIS_2$NEW_appops_status, DIS_MV_lender_MIS_2$Remark)


DIS_MV_lender_MIS_2$Upload_Remarks <- str_c(DIS_MV_lender_MIS_2$NEW_appops_description,' ',DIS_MV_lender_MIS_2$loan_amount)


colnames(DIS_MV_lender_MIS_2)[which(names(DIS_MV_lender_MIS_2) == "offer_application_number")] <- "Application_Number"

DIS_MV_lender_MIS_2<-DIS_MV_lender_MIS_2 %>% mutate('Lender'="Money View")

write.xlsx(DIS_MV_lender_MIS_2, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Disbursed_FB_File_2.xlsx")

MV_upload_1<-DIS_MV_lender_MIS_2 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))


MV_upload_1<- MV_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=MV_upload_1$loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



write.xlsx(MV_upload_1, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_1_2.xlsx")


######################



Status_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV2.xlsx", sheet = "status") %>% select(Mobile,Offer_selected,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved)

offer<-Status_MV_lender_MIS %>% filter(Offer_selected==1) %>% select(Mobile) %>% mutate(`REMARKS`="Offer_selected")


Rejected<-Status_MV_lender_MIS %>% filter(Rejected==1) %>% select(Mobile) %>% mutate(`REMARKS`="Rejected")

Cancelled<-Status_MV_lender_MIS %>% filter(Cancelled==1) %>% select(Mobile) %>% mutate(`REMARKS`="Cancelled")

Submits<-Status_MV_lender_MIS %>% filter(Submits==1) %>% select(Mobile) %>% mutate(`REMARKS`="Submits")

In_Process<-Status_MV_lender_MIS %>% filter(In_Process==1) %>% select(Mobile) %>% mutate(`REMARKS`="In_Process")

NACH<-Status_MV_lender_MIS %>% filter(NACH==1) %>% select(Mobile) %>% mutate(`REMARKS`="NACH")

Loan_Approved<-Status_MV_lender_MIS %>% filter(Loan_Approved==1) %>% select(Mobile) %>% mutate(`REMARKS`="Loan_Approved")


PRO_MV_lender_MIS_2<-rbind(offer,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved)

colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "Mobile")] <- "mobile_number"


colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "REMARKS")] <- "customer_status_name"


#PRO_MV_lender_MIS_2$mobile_number <- as.numeric(PRO_MV_lender_MIS_2$mobile_number)


PRO_MV_lender_MIS_2$ph_no<-DIS_MV_lender_MIS_2$mobile_number[match(PRO_MV_lender_MIS_2$mobile_number, DIS_MV_lender_MIS_2$mobile_number)]
PRO_MV_lender_MIS_2$ph_no[is.na(PRO_MV_lender_MIS_2$ph_no)] <- 'FALSE'

PRO_MV_lender_MIS_2<-PRO_MV_lender_MIS_2 %>% filter(ph_no=='FALSE')

MV_appos_dump$ph_no<-PRO_MV_lender_MIS_2$mobile_number[match(MV_appos_dump$phone_home, PRO_MV_lender_MIS_2$mobile_number)]

PRO_MV_lender_MIS_2$offer_application_number<-MV_appos_dump$offer_application_number[match(PRO_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]



PRO_MV_lender_MIS_2$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(PRO_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]

PRO_MV_lender_MIS_2$NEW_appops_status<-appos_map_MV$New_Status[match(PRO_MV_lender_MIS_2$customer_status_name, appos_map_MV$CM_Status)]

PRO_MV_lender_MIS_2$NEW_appops_description<-appos_map_MV$Status_Description[match(PRO_MV_lender_MIS_2$customer_status_name, appos_map_MV$CM_Status)]


PRO_MV_lender_MIS_2 <-PRO_MV_lender_MIS_2 %>%
  mutate(Remark = case_when(
    is.na(offer_application_number) ~ 'Not in CRM',
    NEW_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (NEW_appops_status ==990) ~ '990',
    (NEW_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (NEW_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (NEW_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

PRO_MV_lender_MIS_2$Remark <- ifelse(is.na(PRO_MV_lender_MIS_2$Remark), PRO_MV_lender_MIS_2$NEW_appops_status, PRO_MV_lender_MIS_2$Remark)



PRO_MV_lender_MIS_2$Upload_Remarks <- str_c(PRO_MV_lender_MIS_2$customer_status_name,'-',PRO_MV_lender_MIS_2$NEW_appops_description)

colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "offer_application_number")] <- "Application_Number"


PRO_MV_lender_MIS_2<-PRO_MV_lender_MIS_2 %>% mutate('Lender'="Money View")

write.xlsx(PRO_MV_lender_MIS_2, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Proposal_FB_File_2.xlsx") 

MV_upload_2<-PRO_MV_lender_MIS_2 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

MV_upload_2<- MV_upload_2 %>% filter(!grepl('Repeated_Feedback_cases, Not_in_CRM',Remark,ignore.case=TRUE)) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(MV_upload_2, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_2_2.xlsx")

