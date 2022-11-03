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

RBL_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\RBL_CC.xlsx") %>% select(offer_reference_number,customer_status_name)

#view(RBL_lender_MIS)
appos_map_RBL <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'RBL CC')
RBL_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('RBL',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_reference_number,offer_application_number,status,name,appops_status_code)
RBL_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('RBL',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_reference_number,offer_application_number,status,name,appops_status_code)
RBL_appos_dump<-rbind(RBL_appos_dump1,RBL_appos_dump2)

########################RBL LenderFeedBack


RBL_lender_MIS$offer_application_number<-RBL_appos_dump$offer_application_number[match(RBL_lender_MIS$offer_reference_number, RBL_appos_dump$offer_reference_number)]

RBL_lender_MIS$mobile_number<-RBL_appos_dump$phone_home[match(RBL_lender_MIS$offer_application_number, RBL_appos_dump$offer_application_number)]


#RBL_lender_MIS$offer_application_number<-RBL_appos_dump$offer_application_number[match(RBL_lender_MIS$leadid, RBL_appos_dump$lead_id)]



RBL_lender_MIS$CURRENT_appops_status<-RBL_appos_dump$appops_status_code[match(RBL_lender_MIS$offer_application_number, RBL_appos_dump$offer_application_number)]


RBL_lender_MIS$New_appops_status<-appos_map_RBL$New_Status[match(RBL_lender_MIS$customer_status_name, appos_map_RBL$CM_Status)]

RBL_lender_MIS$NEW_appops_description<-appos_map_RBL$Status_Description[match(RBL_lender_MIS$customer_status_name, appos_map_RBL$CM_Status)]


RBL_lender_MIS <-RBL_lender_MIS %>%
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

RBL_lender_MIS$Remark <- ifelse(is.na(RBL_lender_MIS$Remark), RBL_lender_MIS$New_appops_status, RBL_lender_MIS$Remark)

#RBL_lender_MIS<-RBL_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
RBL_lender_MIS$Upload_Remarks <- paste(RBL_lender_MIS$NEW_appops_description,'-',RBL_lender_MIS$customer_status_name)

RBL_lender_MIS$Remark[RBL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',RBL_lender_MIS$New_appops_status,ignore.case=TRUE) & RBL_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, RBL_lender_MIS$New_appops_status %in% c(280,380) 

RBL_lender_MIS$Remark[RBL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',RBL_lender_MIS$New_appops_status,ignore.case=TRUE) & RBL_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, RBL_lender_MIS$New_appops_status %in% c(480) 

RBL_lender_MIS$Remark[RBL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',RBL_lender_MIS$New_appops_status,ignore.case=TRUE) & RBL_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, RBL_lender_MIS$New_appops_status %in% c(580) 

RBL_lender_MIS$Remark[RBL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',RBL_lender_MIS$New_appops_status,ignore.case=TRUE) & RBL_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, RBL_lender_MIS$New_appops_status %in% c(680) 

colnames(RBL_lender_MIS)[which(names(RBL_lender_MIS) == "offer_application_number")] <- "Application_Number"

RBL_lender_MIS<-RBL_lender_MIS %>% mutate('Lender'="RBL",`loan_amount`='',`Customer_requested_loan_amount`='')


write.xlsx(RBL_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_RBL_FB_File.xlsx")

RBL_upload<-RBL_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

RBL_upload<- RBL_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=RBL_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=RBL_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(RBL_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\RBL_appos_dump.xlsx")


#################################
