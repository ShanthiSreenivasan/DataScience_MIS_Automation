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

LNT_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\LNT.xlsx") %>% select(Applicant.mobile.number,LoginDate,stageStatus,totalLoanAmount,DISBURSALDATE)

appos_map_LNT <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'LNT')


LNT_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('L&T Finance',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
LNT_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('L&T Finance',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)

LNT_appos_dump<-rbind(LNT_appos_dump1,LNT_appos_dump2)
#View(LNT_lender_MIS)
#LNT_appos_dump<-replace(LNT_appos_dump$names, "L&T Finance", "LNT")
# %>% filter(grepl('L&T',names,ignore.case=TRUE)) 
########################LNT LenderFeedBack


colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "stageStatus")] <- "customer_status_name"

colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "Applicant.mobile.number")] <- "mobile_number"

LNT_lender_MIS$mobile_number <- as.numeric(LNT_lender_MIS$mobile_number)


LNT_lender_MIS$ph_no<-LNT_appos_dump$phone_home[match(LNT_lender_MIS$mobile_number,LNT_appos_dump$phone_home)]


LNT_lender_MIS <-LNT_lender_MIS %>%
  mutate(leads_flag = case_when(
    !is.na(LNT_lender_MIS$ph_no) ~ 'TRUE'

  ))

LNT_lender_MIS$leads_flag[is.na(LNT_lender_MIS$leads_flag)] <- 'FALSE'

LNT_lender_MIS<-LNT_lender_MIS %>% filter(leads_flag=='TRUE')

LNT_lender_MIS$offer_application_number<-LNT_appos_dump$offer_application_number[match(LNT_lender_MIS$mobile_number, LNT_appos_dump$phone_home)]



LNT_lender_MIS$CURRENT_appops_status<-LNT_appos_dump$appops_status_code[match(LNT_lender_MIS$mobile_number, LNT_appos_dump$phone_home)]


LNT_lender_MIS$New_appops_status<-appos_map_LNT$New_Status[match(LNT_lender_MIS$customer_status_name, appos_map_LNT$CM_Status)]

LNT_lender_MIS$NEW_appops_description<-appos_map_LNT$Status_Description[match(LNT_lender_MIS$customer_status_name, appos_map_LNT$CM_Status)]



LNT_lender_MIS <-LNT_lender_MIS %>%
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

LNT_lender_MIS$Remark <- ifelse(is.na(LNT_lender_MIS$Remark), LNT_lender_MIS$New_appops_status, LNT_lender_MIS$Remark)


LNT_lender_MIS$Upload_Remarks <- paste(LNT_lender_MIS$customer_status_name,'-',LNT_lender_MIS$NEW_appops_description)


LNT_lender_MIS$Remark[LNT_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LNT_lender_MIS$New_appops_status,ignore.case=TRUE) & LNT_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, LNT_lender_MIS$New_appops_status %in% c(280,380) 

LNT_lender_MIS$Remark[LNT_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LNT_lender_MIS$New_appops_status,ignore.case=TRUE) & LNT_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, LNT_lender_MIS$New_appops_status %in% c(480) 

LNT_lender_MIS$Remark[LNT_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LNT_lender_MIS$New_appops_status,ignore.case=TRUE) & LNT_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, LNT_lender_MIS$New_appops_status %in% c(580) 

LNT_lender_MIS$Remark[LNT_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',LNT_lender_MIS$New_appops_status,ignore.case=TRUE) & LNT_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, LNT_lender_MIS$New_appops_status %in% c(680) 

colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(LNT_lender_MIS)[which(names(LNT_lender_MIS) == "totalLoanAmount")] <- "loan_amount"



LNT_lender_MIS<-LNT_lender_MIS %>% mutate('Lender'="L&T Finance",`Customer_requested_loan_amount`='')

LNT_upload<-LNT_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

LNT_upload<- LNT_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=LNT_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=LNT_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

write.xlsx(LNT_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_LNT_FB_File.xlsx")

#write.xlsx(LNT_upload, file = "C:\\R\\Lender Feedback\\Output\\LNT_upload.xlsx")

#################################
