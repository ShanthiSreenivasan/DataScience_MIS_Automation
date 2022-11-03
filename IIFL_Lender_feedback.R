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

IIFL_lender_MIS<- read.csv("C:\\R\\Lender Feedback\\Input\\IIFL.csv") #%>% select(mobile,status,loan_status,bank_statement_status,loan_application_no,bank_statement_failure_reason)




#names(IIFL_lender_MIS)
appos_map_IIFL <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'IIFL')

IIFL_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('IIFL',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
IIFL_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('IIFL',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
IIFL_appos_dump<-rbind(IIFL_appos_dump1,IIFL_appos_dump2)
########################IIFL LenderFeedBack

IIFL_lender_MIS$mobile <- as.numeric(IIFL_lender_MIS$mobile)

colnames(IIFL_lender_MIS)[which(names(IIFL_lender_MIS) == "mobile")] <- "mobile_number"

colnames(IIFL_lender_MIS)[which(names(IIFL_lender_MIS) == "status")] <- "customer_status_name"

#IIFL_appos_dump$ph_no<-IIFL_lender_MIS$mobile_number[match(IIFL_appos_dump$phone_home, IIFL_lender_MIS$mobile_number)]

IIFL_lender_MIS$offer_application_number<-IIFL_appos_dump$offer_application_number[match(IIFL_lender_MIS$mobile_number, IIFL_appos_dump$phone_home)]



IIFL_lender_MIS$CURRENT_appops_status<-IIFL_appos_dump$appops_status_code[match(IIFL_lender_MIS$mobile_number, IIFL_appos_dump$phone_home)]

IIFL_lender_MIS$New_appops_status<-appos_map_IIFL$New_Status[match(IIFL_lender_MIS$customer_status_name, appos_map_IIFL$CM_Status)]

IIFL_lender_MIS$NEW_appops_description<-appos_map_IIFL$Status_Description[match(IIFL_lender_MIS$customer_status_name, appos_map_IIFL$CM_Status)]


IIFL_lender_MIS <-IIFL_lender_MIS %>%
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

IIFL_lender_MIS$Remark <- ifelse(is.na(IIFL_lender_MIS$Remark), IIFL_lender_MIS$New_appops_status, IIFL_lender_MIS$Remark)

IIFL_lender_MIS$Upload_Remarks <- paste(IIFL_lender_MIS$customer_status_name,"-",IIFL_lender_MIS$NEW_appops_description,"-",IIFL_lender_MIS$loan_status)

IIFL_lender_MIS$Remark[IIFL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IIFL_lender_MIS$New_appops_status,ignore.case=TRUE) & IIFL_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, IIFL_lender_MIS$New_appops_status %in% c(280,380) 

IIFL_lender_MIS$Remark[IIFL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IIFL_lender_MIS$New_appops_status,ignore.case=TRUE) & IIFL_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, IIFL_lender_MIS$New_appops_status %in% c(480) 

IIFL_lender_MIS$Remark[IIFL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IIFL_lender_MIS$New_appops_status,ignore.case=TRUE) & IIFL_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, IIFL_lender_MIS$New_appops_status %in% c(580) 

IIFL_lender_MIS$Remark[IIFL_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',IIFL_lender_MIS$New_appops_status,ignore.case=TRUE) & IIFL_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, IIFL_lender_MIS$New_appops_status %in% c(680) 

colnames(IIFL_lender_MIS)[which(names(IIFL_lender_MIS) == "offer_application_number")] <- "Application_Number"

#colnames(IIFL_lender_MIS)[which(names(IIFL_lender_MIS) == "offer_application_number")] <- "Application_Number"

IIFL_lender_MIS<-IIFL_lender_MIS %>% mutate('Lender'="IIFL", `loan_amount`='',`Customer_requested_loan_amount`='')


write.xlsx(IIFL_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_IIFL_FB_File.xlsx")

IIFL_upload_1<-IIFL_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))





IIFL_upload_1<- IIFL_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=IIFL_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=IIFL_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(IIFL_upload_1, file = "C:\\R\\Lender Feedback\\Output\\IIFL_upload.xlsx")

