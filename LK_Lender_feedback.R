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

LK_lender_MIS <- read.csv("C:\\R\\Lender Feedback\\Input\\LK.csv") %>% select(Lead.Id,Mobile.Number,Application.Id,Status,Substatus,Rejection.Reason)

#names(LK_lender_MIS)
appos_map_LK <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Lendingkart')


LK_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Lending',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
LK_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Lending',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

LK_appos_dump<-rbind(LK_appos_dump1,LK_appos_dump2)
########################LK LenderFeedBack

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "Mobile.Number")] <- "mobile_number"

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "Substatus")] <- "customer_status_name"

LK_lender_MIS$mobile_number <- as.numeric(LK_lender_MIS$mobile_number)


LK_lender_MIS$offer_application_number<-LK_appos_dump$offer_application_number[match(LK_lender_MIS$mobile_number, LK_appos_dump$phone_home)]



LK_lender_MIS$CURRENT_appops_status<-LK_appos_dump$appops_status_code[match(LK_lender_MIS$mobile_number, LK_appos_dump$phone_home)]


LK_lender_MIS$New_appops_status<-appos_map_LK$New_Status[match(LK_lender_MIS$customer_status_name, appos_map_LK$CM_Status)]

LK_lender_MIS$NEW_appops_description<-appos_map_LK$Status_Description[match(LK_lender_MIS$customer_status_name, appos_map_LK$CM_Status)]


LK_lender_MIS <-LK_lender_MIS %>%
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

LK_lender_MIS$Remark <- ifelse(is.na(LK_lender_MIS$Remark), LK_lender_MIS$New_appops_status, LK_lender_MIS$Remark)

LK_lender_MIS$Upload_Remarks <- paste(LK_lender_MIS$Status,'-',LK_lender_MIS$customer_status_name)


LK_lender_MIS$Remark[LK_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LK_lender_MIS$New_appops_status,ignore.case=TRUE) & LK_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, LK_lender_MIS$New_appops_status %in% c(280,380) 

LK_lender_MIS$Remark[LK_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LK_lender_MIS$New_appops_status,ignore.case=TRUE) & LK_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, LK_lender_MIS$New_appops_status %in% c(480) 

LK_lender_MIS$Remark[LK_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LK_lender_MIS$New_appops_status,ignore.case=TRUE) & LK_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, LK_lender_MIS$New_appops_status %in% c(580) 

LK_lender_MIS$Remark[LK_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LK_lender_MIS$New_appops_status,ignore.case=TRUE) & LK_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, LK_lender_MIS$New_appops_status %in% c(680) 

colnames(LK_lender_MIS)[which(names(LK_lender_MIS) == "offer_application_number")] <- "Application_Number"


LK_lender_MIS<-LK_lender_MIS %>% mutate('Lender'="Lendingkart",`loan_amount`='',`Customer_requested_loan_amount`='')

write.xlsx(LK_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_LK_FB_File.xlsx")

LK_upload<-LK_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

LK_upload<- LK_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=LK_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=LK_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(LK_upload, file = "C:\\R\\Lender Feedback\\Output\\LK_upload.xlsx")

#################################
