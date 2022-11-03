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
#library(yaml)
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

IDFC_lender_MIS1<- read.xlsx("C:\\R\\Lender Feedback\\Input\\IDFC_CC.xlsx", sheet = "Approved master") %>% select(CRM.Lead.Id,Applicant.Mobile,Stage,Sub.Stage,Stage.by.Integration.Status,FI.Curing.Reason)

IDFC_lender_MIS2<- read.xlsx("C:\\R\\Lender Feedback\\Input\\IDFC_CC.xlsx", sheet = "Secure card Master") %>% select(CRM.Lead.Id,Applicant.Mobile,Stage,Sub.Stage,Stage.by.Integration.Status,FI.Curing.Reason)

IDFC_lender_MIS<-rbind(IDFC_lender_MIS1,IDFC_lender_MIS2)

appos_map_IDFC <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'IDFC CC')

IDFC_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('IDFC',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
IDFC_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('IDFC',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
IDFC_appos_dump<-rbind(IDFC_appos_dump1,IDFC_appos_dump2)

########################IDFC LenderFeedBack

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "Applicant.Mobile")] <- "mobile_number"

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "Sub.Stage")] <- "customer_status_name"

#colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "Stage.by.Integration.Status")] <- "customer_status_name"


IDFC_lender_MIS$mobile_number <- as.numeric(IDFC_lender_MIS$mobile_number)


IDFC_lender_MIS$offer_application_number<-IDFC_appos_dump$offer_application_number[match(IDFC_lender_MIS$mobile_number, IDFC_appos_dump$phone_home)]



IDFC_lender_MIS$CURRENT_appops_status<-IDFC_appos_dump$appops_status_code[match(IDFC_lender_MIS$mobile_number, IDFC_appos_dump$phone_home)]


IDFC_lender_MIS$New_appops_status<-appos_map_IDFC$New_Status[match(IDFC_lender_MIS$customer_status_name, appos_map_IDFC$CM_Status)]

IDFC_lender_MIS$NEW_appops_description<-appos_map_IDFC$Status_Description[match(IDFC_lender_MIS$customer_status_name, appos_map_IDFC$CM_Status)]

names(IDFC_lender_MIS)
IDFC_lender_MIS <-IDFC_lender_MIS %>%
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

IDFC_lender_MIS$Remark <- ifelse(is.na(IDFC_lender_MIS$Remark), IDFC_lender_MIS$New_appops_status, IDFC_lender_MIS$Remark)

#IDFC_lender_MIS<-IDFC_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
IDFC_lender_MIS$Upload_Remarks <- paste(IDFC_lender_MIS$customer_status_name,'-',IDFC_lender_MIS$Stage.by.Integration.Status)


IDFC_lender_MIS$Remark[IDFC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IDFC_lender_MIS$New_appops_status,ignore.case=TRUE) & IDFC_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, IDFC_lender_MIS$New_appops_status %in% c(280,380) 

IDFC_lender_MIS$Remark[IDFC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IDFC_lender_MIS$New_appops_status,ignore.case=TRUE) & IDFC_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, IDFC_lender_MIS$New_appops_status %in% c(480) 

IDFC_lender_MIS$Remark[IDFC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IDFC_lender_MIS$New_appops_status,ignore.case=TRUE) & IDFC_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, IDFC_lender_MIS$New_appops_status %in% c(580) 

IDFC_lender_MIS$Remark[IDFC_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IDFC_lender_MIS$New_appops_status,ignore.case=TRUE) & IDFC_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, IDFC_lender_MIS$New_appops_status %in% c(680) 

colnames(IDFC_lender_MIS)[which(names(IDFC_lender_MIS) == "offer_application_number")] <- "Application_Number"


IDFC_lender_MIS<-IDFC_lender_MIS %>% mutate('Lender'="IDFC", `loan_amount`='',`Customer_requested_loan_amount`='')

write.xlsx(IDFC_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_IDFC_FB_File.xlsx")

IDFC_upload<-IDFC_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

IDFC_upload_1<-IDFC_upload %>% filter(Remark %in% c(990))

IDFC_upload_1<- IDFC_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=IDFC_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=IDFC_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`='75000',`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(IDFC_upload_1, file = "C:\\R\\Lender Feedback\\Output\\IDFC_upload_1.xlsx")


#IDFC_upload_2<-IDFC_upload %>% filter(!Remark %in% c(990))

# IDFC_upload_2<- IDFC_upload_2 %>% select(Application_Number) %>% 
#   mutate(`App_Ops_Status`=IDFC_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=IDFC_upload_2$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")
# 

#write.xlsx(IDFC_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\IDFC_appos_dump.xlsx")

#################################
