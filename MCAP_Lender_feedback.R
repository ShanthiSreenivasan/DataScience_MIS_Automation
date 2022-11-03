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
#library(logging)
#library(mailR)
library(xtable)
library(yaml)
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

MCAP_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\MCAP.xlsx", sheet='DATA') %>% select(CustomerPhoneNumber,Calling.Disposition,Final.Disposition,Final.Loan.Amount)




#names(MCAP_lender_MIS)
appos_map_MCAP <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'MCAP')


MCAP_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('M Capital',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
MCAP_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('M Capital',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
MCAP_appos_dump<-rbind(MCAP_appos_dump1,MCAP_appos_dump2)
########################MCAP LenderFeedBack

MCAP_lender_MIS$CustomerPhoneNumber <- as.numeric(MCAP_lender_MIS$CustomerPhoneNumber)

colnames(MCAP_lender_MIS)[which(names(MCAP_lender_MIS) == "CustomerPhoneNumber")] <- "mobile_number"

colnames(MCAP_lender_MIS)[which(names(MCAP_lender_MIS) == "Calling.Disposition")] <- "customer_status_name"

#MCAP_appos_dump$ph_no<-MCAP_lender_MIS$mobile_number[match(MCAP_appos_dump$phone_home, MCAP_lender_MIS$mobile_number)]

MCAP_lender_MIS$offer_application_number<-MCAP_appos_dump$offer_application_number[match(MCAP_lender_MIS$mobile_number, MCAP_appos_dump$phone_home)]



MCAP_lender_MIS$CURRENT_appops_status<-MCAP_appos_dump$appops_status_code[match(MCAP_lender_MIS$mobile_number, MCAP_appos_dump$phone_home)]

MCAP_lender_MIS$New_appops_status<-appos_map_MCAP$New_Status[match(MCAP_lender_MIS$customer_status_name, appos_map_MCAP$CM_Status)]

MCAP_lender_MIS$NEW_appops_description<-appos_map_MCAP$Status_Description[match(MCAP_lender_MIS$customer_status_name, appos_map_MCAP$CM_Status)]


MCAP_lender_MIS <-MCAP_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) & !is.na(offer_application_number)  ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

MCAP_lender_MIS$Remark <- ifelse(is.na(MCAP_lender_MIS$Remark), MCAP_lender_MIS$New_appops_status, MCAP_lender_MIS$Remark)

MCAP_lender_MIS$Upload_Remarks <- paste(MCAP_lender_MIS$Final.Disposition,"-",MCAP_lender_MIS$Remark)

MCAP_lender_MIS$Remark[MCAP_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',MCAP_lender_MIS$New_appops_status,ignore.case=TRUE) & MCAP_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, MCAP_lender_MIS$New_appops_status %in% c(280,380) 

MCAP_lender_MIS$Remark[MCAP_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',MCAP_lender_MIS$New_appops_status,ignore.case=TRUE) & MCAP_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, MCAP_lender_MIS$New_appops_status %in% c(480) 

MCAP_lender_MIS$Remark[MCAP_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',MCAP_lender_MIS$New_appops_status,ignore.case=TRUE) & MCAP_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, MCAP_lender_MIS$New_appops_status %in% c(580) 

MCAP_lender_MIS$Remark[MCAP_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',MCAP_lender_MIS$New_appops_status,ignore.case=TRUE) & MCAP_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, MCAP_lender_MIS$New_appops_status %in% c(680) 


colnames(MCAP_lender_MIS)[which(names(MCAP_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(MCAP_lender_MIS)[which(names(MCAP_lender_MIS) == "Final.Loan.Amount")] <- "loan_amount"



MCAP_lender_MIS<-MCAP_lender_MIS %>% mutate('Lender'="M Capital",`Customer_requested_loan_amount`='')


write.xlsx(MCAP_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_MCAP_FB_File.xlsx")

MCAP_upload_1<-MCAP_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))





MCAP_upload_1<- MCAP_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MCAP_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MCAP_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(MCAP_upload_1, file = "C:\\R\\Lender Feedback\\Output\\MCAP_upload.xlsx")

