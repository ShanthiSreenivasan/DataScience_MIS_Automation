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

CITI_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CITI_CC.xlsx") %>% select(Creative,V_APPLICATION_ID,final_status,reject.reasons)

#view(CITI_lender_MIS)
appos_map_CITI <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'CITICC')
CITI_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('CITI',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
CITI_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('CITI',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
CITI_appos_dump<-rbind(CITI_appos_dump1,CITI_appos_dump2)

########################CITI LenderFeedBack

#colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "Creative")] <- "leadid"

colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "final_status")] <- "customer_status_name"

CITI_lender_MIS$leadid <-as.integer(sub("^00", "", CITI_lender_MIS$Creative))


#View(CITI_lender_MIS)
#names(CITI_lender_MIS)
CITI_lender_MIS$mobile_number<-CITI_appos_dump$phone_home[match(CITI_lender_MIS$leadid, CITI_appos_dump$lead_id)]


CITI_lender_MIS$offer_application_number<-CITI_appos_dump$offer_application_number[match(CITI_lender_MIS$leadid, CITI_appos_dump$lead_id)]



CITI_lender_MIS$CURRENT_appops_status<-CITI_appos_dump$appops_status_code[match(CITI_lender_MIS$leadid, CITI_appos_dump$lead_id)]


CITI_lender_MIS$New_appops_status<-appos_map_CITI$New_Status[match(CITI_lender_MIS$customer_status_name, appos_map_CITI$CM_Status)]

CITI_lender_MIS$NEW_appops_description<-appos_map_CITI$Status_Description[match(CITI_lender_MIS$customer_status_name, appos_map_CITI$CM_Status)]


CITI_lender_MIS <-CITI_lender_MIS %>%
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

CITI_lender_MIS$Remark <- ifelse(is.na(CITI_lender_MIS$Remark), CITI_lender_MIS$New_appops_status, CITI_lender_MIS$Remark)

#CITI_lender_MIS<-CITI_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
CITI_lender_MIS$Upload_Remarks <- paste(CITI_lender_MIS$NEW_appops_description,'-',CITI_lender_MIS$customer_status_name)


CITI_lender_MIS$Remark[CITI_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CITI_lender_MIS$New_appops_status,ignore.case=TRUE) & CITI_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, CITI_lender_MIS$New_appops_status %in% c(280,380) 

CITI_lender_MIS$Remark[CITI_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CITI_lender_MIS$New_appops_status,ignore.case=TRUE) & CITI_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, CITI_lender_MIS$New_appops_status %in% c(480) 

CITI_lender_MIS$Remark[CITI_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CITI_lender_MIS$New_appops_status,ignore.case=TRUE) & CITI_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, CITI_lender_MIS$New_appops_status %in% c(580) 

CITI_lender_MIS$Remark[CITI_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',CITI_lender_MIS$New_appops_status,ignore.case=TRUE) & CITI_lender_MIS$CURRENT_appops_status %in% c(650,660,670,690,700,701,780)] <- 680#, CITI_lender_MIS$New_appops_status %in% c(680) 

colnames(CITI_lender_MIS)[which(names(CITI_lender_MIS) == "offer_application_number")] <- "Application_Number"

CITI_lender_MIS<-CITI_lender_MIS %>% mutate('Lender'="CITI",`loan_amount`='',`Customer_requested_loan_amount`='')


write.xlsx(CITI_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_CITI_FB_File.xlsx")

CITI_upload<-CITI_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

CITI_upload<- CITI_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=CITI_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=CITI_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(CITI_upload, file = "C:\\R\\Lender Feedback\\Output\\CITI_upload.xlsx")

#################################
