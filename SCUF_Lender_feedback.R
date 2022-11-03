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


DIS_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Lead_Disbursed_Data_Report.xlsx", sheet='file') %>% select(CMIDENTIFIER,PRPSLNO,PERMOBILE,LNAMT)



PRO_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Lead_Data_Proposal_Report.xlsx", sheet='file') %>% select(CMIDENTIFIER,MOBILENO,REMARKS,REASONFORREJECTION,USERREMARKS,LOANAMOUNT)


#DIS_lender_MIS
appos_map_SCUF <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'sriram1')


SCUF_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Shriram',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
SCUF_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Shriram',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
SCUF_appos_dump<-rbind(SCUF_appos_dump1,SCUF_appos_dump2)


########################SCUF LenderFeedBack
colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "PERMOBILE")] <- "mobile_number"

DIS_lender_MIS$lead_id <-as.integer(sub("^CM", "", DIS_lender_MIS$CMIDENTIFIER))
SCUF_appos_dump$ph_no<-DIS_lender_MIS$mobile_number[match(SCUF_appos_dump$phone_home, DIS_lender_MIS$mobile_number)]

DIS_lender_MIS$offer_application_number<-SCUF_appos_dump$offer_application_number[match(DIS_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]



DIS_lender_MIS$CURRENT_appops_status<-SCUF_appos_dump$appops_status_code[match(DIS_lender_MIS$mobile_number, SCUF_appos_dump$phone_home)]


DIS_lender_MIS$customer_status_name<-'Loan Disbursed'

DIS_lender_MIS$New_appops_status<- '990'

DIS_lender_MIS$NEW_appops_description<-"Loan Disbursed"


DIS_lender_MIS <-DIS_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    is.na(offer_application_number) ~ 'Not_in_CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990'
    
  ))

DIS_lender_MIS$Remark <- ifelse(is.na(DIS_lender_MIS$Remark), DIS_lender_MIS$New_appops_status, DIS_lender_MIS$Remark)


#DIS_lender_MIS$NEW_appops_description<-appos_map_SCUF$Description[match(DIS_lender_MIS$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]

DIS_lender_MIS$Upload_Remarks <- str_c(DIS_lender_MIS$PRPSLNO,' ',DIS_lender_MIS$NEW_appops_description,' ',DIS_lender_MIS$LNAMT)

#


colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(DIS_lender_MIS)[which(names(DIS_lender_MIS) == "LNAMT")] <- "loan_amount"


DIS_lender_MIS<-DIS_lender_MIS %>% mutate('Lender'="Shriram",`Customer_requested_loan_amount`='')

write.xlsx(DIS_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SCUF_Disbursed_FB_File.xlsx")


# SCUF_upload_1<- DIS_lender_MIS %>% filter(!NEW_appops_description %in% c("Repeated_Feedback_cases", "Not_in_CRM")) %>% select(Application_Number) %>% 
#   mutate(`App_Ops_Status`=DIS_lender_MIS$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=DIS_lender_MIS$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(SCUF_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_1.xlsx")


scuf_df1<-DIS_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

SCUF_upload_1<- scuf_df1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=scuf_df1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=scuf_df1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=scuf_df1$LNAMT,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(SCUF_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_1.xlsx")

######################

PRO_lender_MIS$lead_id <-as.integer(sub("^CM", "", PRO_lender_MIS$CMIDENTIFIER))

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
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) & !is.na(offer_application_number) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

PRO_lender_MIS$Remark <- ifelse(is.na(PRO_lender_MIS$Remark), PRO_lender_MIS$New_appops_status, PRO_lender_MIS$Remark)

PRO_lender_MIS$Remark[is.na(PRO_lender_MIS$Remark)] <- "Repeated_Feedback_cases"


PRO_lender_MIS$Upload_Remarks <- str_c(PRO_lender_MIS$customer_status_name,'-',PRO_lender_MIS$REASONFORREJECTION,'-',PRO_lender_MIS$USERREMARKS)

PRO_lender_MIS$Upload_Remarks <- ifelse(is.na(PRO_lender_MIS$Upload_Remarks), PRO_lender_MIS$customer_status_name, PRO_lender_MIS$Upload_Remarks)


PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, PRO_lender_MIS$New_appops_status %in% c(280,380) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, PRO_lender_MIS$New_appops_status %in% c(480) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, PRO_lender_MIS$New_appops_status %in% c(580) 

PRO_lender_MIS$Remark[PRO_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, PRO_lender_MIS$New_appops_status %in% c(680) 

colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "offer_application_number")] <- "Application_Number"


colnames(PRO_lender_MIS)[which(names(PRO_lender_MIS) == "LOANAMOUNT")] <- "loan_amount"

PRO_lender_MIS<-PRO_lender_MIS %>% mutate('Lender'="Shriram", `Customer_requested_loan_amount`='')

write.xlsx(PRO_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SCUF_Proposal_FB_File.xlsx")

scuf_df2<-PRO_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

SCUF_upload_2<- scuf_df2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=scuf_df2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=scuf_df2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(SCUF_upload_2, file = "C:\\R\\Lender Feedback\\Output\\SCUF_upload_2.xlsx")

PRO_chk<-PRO_lender_MIS %>% filter(Remark %in% c('Not_in_CRM')) %>% select(CMIDENTIFIER,mobile_number,loan_amount,lead_id,customer_status_name,Lender,Customer_requested_loan_amount,Application_Number,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)


DIS_chk<-DIS_lender_MIS %>% filter(Remark %in% c('Not_in_CRM')) %>% select(CMIDENTIFIER,mobile_number,loan_amount,lead_id,customer_status_name,Lender,Customer_requested_loan_amount,Application_Number,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)


#DIS_lender_MIS<-DIS_lender_MIS %>% select(CMIDENTIFIER,mobile_number,lead_id,Application_Number,customer_status_name,REASONFORREJECTION,USERREMARKS,loan_amount,Lender)
CRM_na<-rbind(PRO_chk,DIS_chk)

CRM_na<-CRM_na %>% select(CMIDENTIFIER,mobile_number,loan_amount,lead_id,customer_status_name,Lender,Customer_requested_loan_amount)

CRM_na$ph_no<-SCUF_appos_dump$offer_application_number[match(CRM_na$lead_id,SCUF_appos_dump$lead_id)]

colnames(CRM_na)[which(names(CRM_na) == "ph_no")] <- "Application_Number"


CRM_na$CURRENT_appops_status<-SCUF_appos_dump$appops_status_code[match(CRM_na$lead_id, SCUF_appos_dump$lead_id)]

CRM_na$New_appops_status<-appos_map_SCUF$New_Status[match(CRM_na$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]

CRM_na$NEW_appops_description<-appos_map_SCUF$Status_Description[match(CRM_na$customer_status_name, appos_map_SCUF$Campaign_Sub_Status)]




#colnames(CRM_na)[which(names(CRM_na) == "Remark")] <- "pre_Remark"


CRM_na <-CRM_na %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) & !is.na(Application_Number) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(Application_Number) ~ 'Not_in_CRM'
    
    
  ))

CRM_na$Remark <- ifelse(is.na(CRM_na$Remark), CRM_na$New_appops_status, CRM_na$Remark)

CRM_na$Remark[is.na(CRM_na$Remark)] <- "Repeated_Feedback_cases"

CRM_na$Upload_Remarks <- str_c(CRM_na$customer_status_name,'-',CRM_na$NEW_appops_description)

CRM_na$Upload_Remarks <- ifelse(is.na(CRM_na$Upload_Remarks), CRM_na$customer_status_name, CRM_na$Upload_Remarks)

CRM_na$Remark[CRM_na$Remark == 'Repeated_Feedback_cases' & grepl('80',CRM_na$New_appops_status,ignore.case=TRUE) & CRM_na$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, CRM_na$New_appops_status %in% c(280,380) 

CRM_na$Remark[CRM_na$Remark == 'Repeated_Feedback_cases' & grepl('80',CRM_na$New_appops_status,ignore.case=TRUE) & CRM_na$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, CRM_na$New_appops_status %in% c(480) 

CRM_na$Remark[CRM_na$Remark == 'Repeated_Feedback_cases' & grepl('80',CRM_na$New_appops_status,ignore.case=TRUE) & CRM_na$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, CRM_na$New_appops_status %in% c(580) 

CRM_na$Remark[CRM_na$Remark == 'Repeated_Feedback_cases' & grepl('80',CRM_na$New_appops_status,ignore.case=TRUE) & CRM_na$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, CRM_na$New_appops_status %in% c(680) 

write.xlsx(CRM_na, file = "C:\\R\\Lender Feedback\\Output\\Revised_SCUF_na.xlsx")

#write.xlsx(SCUF_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\SCUF_appos_dump.xlsx")
