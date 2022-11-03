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

DIS_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\AXIS_PL.xlsx", sheet='Disb') %>% select(APPLICATION_ID,DISB_AMT_AS_ON)

LOGIN_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\AXIS_PL.xlsx", sheet='Login') %>% select(LEAD_ID,APPLICATION_ID)

PRO_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\AXIS_PL.xlsx", sheet='Lead') %>% select(Lead.Id,Mobile,Sub.Product,Status.Code,Reject.Reason,Follow.Up.DT.1)


#names(PRO_lender_MIS)
appos_map_Axis <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Axis')


Axis_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('AXIS Bank',name,ignore.case=TRUE), grepl('PL|SBPL',status,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
Axis_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('AXIS Bank',name,ignore.case=TRUE), grepl('PL|SBPL',status,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
Axis_appos_dump<-rbind(Axis_appos_dump1,Axis_appos_dump2)

########################Axis LenderFeedBack

#colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "APPLICATION_ID")] <- "leadid"

colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "DISB_AMT_AS_ON")] <- "loan_amt"

colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "Lead.Id")] <- "leadid"

LOGIN_lender_MIS$leadid <-as.integer(sub("^CO", "", LOGIN_lender_MIS$LEAD_ID))

DIS_lender_MIS$leadid<-LOGIN_lender_MIS$leadid[match(DIS_lender_MIS$APPLICATION_ID, LOGIN_lender_MIS$APPLICATION_ID)]

DIS_lender_MIS$mobile_number<-PRO_lender_MIS$Mobile[match(DIS_lender_MIS$leadid, PRO_lender_MIS$LEAD_ID)]

#Axis_appos_dump$ph_no<-DIS_lender_MIS$mobile_number[match(Axis_appos_dump$lead_id, DIS_lender_MIS$leadid)]

DIS_lender_MIS$offer_application_number<-Axis_appos_dump$offer_application_number[match(DIS_lender_MIS$mobile_number, Axis_appos_dump$phone_home)]

#View(DIS_lender_MIS)

DIS_lender_MIS$CURRENT_appops_status<-Axis_appos_dump$appops_status_code[match(DIS_lender_MIS$mobile_number, Axis_appos_dump$phone_home)]


DIS_lender_MIS$New_appops_status<- '990'

DIS_lender_MIS$NEW_appops_description<-"Loan Disbursed"


DIS_lender_MIS <-DIS_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not_in_CRM',
    CURRENT_appops_status %in% c(89,140,150,210,240,261,710) ~ 'Stuck_cases',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(offer_application_number) ~ 'Not_in_CRM'
  ))

DIS_lender_MIS$Remark[is.na(DIS_lender_MIS$offer_application_number)] <- 'Not_in_CRM'
DIS_lender_MIS$Remark <- ifelse(is.na(DIS_lender_MIS$Remark), DIS_lender_MIS$New_appops_status, DIS_lender_MIS$Remark)


#DIS_lender_MIS$NEW_appops_description<-appos_map_Axis$Description[match(DIS_lender_MIS$customer_status_name, appos_map_Axis$CM_Status)]

DIS_lender_MIS$Upload_Remarks <- str_c(DIS_lender_MIS$PRPSLNO,' ',DIS_lender_MIS$NEW_appops_description,' ',DIS_lender_MIS$LNAMT)



colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "offer_application_number")] <- "Application_Number"


DIS_lender_MIS<-DIS_lender_MIS %>% mutate('Lender'="AXIS BANK", `Customer_requested_loan_amount`='',`loan_amount`='')

write.xlsx(DIS_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Axis_Disbursed_FB_File.xlsx")

Axis_upload_1<-DIS_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))


Axis_upload_1<- Axis_upload_1 %>% filter(!Remark %in% c('Repeated_Feedback_cases, Not_in_CRM,Stuck_cases')) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Axis_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Axis_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=Axis_upload_1$LNAMT,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(Axis_upload_1, file = "C:\\R\\Lender Feedback\\Output\\Axis_upload_1.xlsx")


######################


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "Mobile")] <- "mobile_number"


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "Status.Code")] <- "customer_status_name"

#PRO_lender_MIS$ph_no<-DIS_lender_MIS$mobile_number[match(PRO_lender_MIS$mobile_number, DIS_lender_MIS$mobile_number)]
#PRO_lender_MIS$ph_no[is.na(PRO_lender_MIS$ph_no)] <- 'FALSE'

#PRO_lender_MIS<-PRO_lender_MIS %>% filter(ph_no=='FALSE')

#Axis_appos_dump$ph_no<-PRO_lender_MIS$mobile_number[match(Axis_appos_dump$phone_home, PRO_lender_MIS$mobile_number)]

PRO_lender_MIS$offer_application_number<-Axis_appos_dump$offer_application_number[match(PRO_lender_MIS$mobile_number, Axis_appos_dump$phone_home)]



PRO_lender_MIS$CURRENT_appops_status<-Axis_appos_dump$appops_status_code[match(PRO_lender_MIS$mobile_number, Axis_appos_dump$phone_home)]

PRO_lender_MIS$New_appops_status<-appos_map_Axis$New_Status[match(PRO_lender_MIS$customer_status_name, appos_map_Axis$CM_Status)]

PRO_lender_MIS$NEW_appops_description<-appos_map_Axis$Status_Description[match(PRO_lender_MIS$customer_status_name, appos_map_Axis$CM_Status)]

#PRO_lender_MIS$NEW_appops_description<-appos_map_Axis$Status_Description[match(PRO_lender_MIS$New_appops_status, appos_map_Axis$New_Status)]

#names(PRO_lender_MIS)

PRO_lender_MIS <-PRO_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

PRO_lender_MIS$Remark <- ifelse(is.na(PRO_lender_MIS$Remark), PRO_lender_MIS$New_appops_status, PRO_lender_MIS$Remark)



PRO_lender_MIS$Upload_Remarks <- str_c(PRO_lender_MIS$customer_status_name,'-',PRO_lender_MIS$Reject.Reason,'-',PRO_lender_MIS$USERREMARKS,'-',PRO_lender_MIS$Follow.Up.DT.1)


PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, PRO_lender_MIS$New_appops_status %in% c(280,380) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, PRO_lender_MIS$New_appops_status %in% c(480) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, PRO_lender_MIS$New_appops_status %in% c(580) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(650,660,670,690,700,701,780)] <- 680#, PRO_lender_MIS$New_appops_status %in% c(680) 


#PRO_lender_MIS<-PRO_lender_MIS %>% filter(!is.na(offer_application_number))

#############to chk


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "offer_application_number")] <- "Application_Number"

PRO_lender_MIS<-PRO_lender_MIS %>% mutate('Lender'="AXIS BANK",`Customer_requested_loan_amount`='',`loan_amount`='')


write.xlsx(PRO_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Axis_Proposal_FB_File.xlsx")

#write.xlsx(Axis_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\Axis_APPOPS.xlsx")

Axis_upload_2<-PRO_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

Axis_upload_2<- Axis_upload_2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Axis_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Axis_upload_2$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

#write.xlsx(Axis_upload_2, file = "C:\\R\\Lender Feedback\\Output\\Axis_upload_2.xlsx")
