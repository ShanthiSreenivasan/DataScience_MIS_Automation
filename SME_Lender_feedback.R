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

SME_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\SME.xlsx", sheet='Sheet0') %>% select(Phone,Partner.Main.Status,Partner.Sub.Status,Approved.Amount,Reject.Remarks,TSE.Comments)




#names(SME_lender_MIS)
appos_map_SME <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'SME')


SME_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Sme Corner',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
SME_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Sme Corner',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
SME_appos_dump<-rbind(SME_appos_dump1,SME_appos_dump2)
########################SME LenderFeedBack

SME_lender_MIS$Phone <- as.numeric(SME_lender_MIS$Phone)

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "Phone")] <- "mobile_number"

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "Partner.Main.Status")] <- "customer_status_name"

#SME_appos_dump$ph_no<-SME_lender_MIS$mobile_number[match(SME_appos_dump$phone_home, SME_lender_MIS$mobile_number)]

SME_lender_MIS$offer_application_number<-SME_appos_dump$offer_application_number[match(SME_lender_MIS$mobile_number, SME_appos_dump$phone_home)]



SME_lender_MIS$CURRENT_appops_status<-SME_appos_dump$appops_status_code[match(SME_lender_MIS$mobile_number, SME_appos_dump$phone_home)]

SME_lender_MIS$New_appops_status<-appos_map_SME$New_Status[match(SME_lender_MIS$customer_status_name, appos_map_SME$CM_Status)]

SME_lender_MIS$NEW_appops_description<-appos_map_SME$Status_Description[match(SME_lender_MIS$customer_status_name, appos_map_SME$CM_Status)]


SME_lender_MIS <-SME_lender_MIS %>%
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

SME_lender_MIS$Remark <- ifelse(is.na(SME_lender_MIS$Remark), SME_lender_MIS$New_appops_status, SME_lender_MIS$Remark)

SME_lender_MIS$Upload_Remarks <- paste(SME_lender_MIS$customer_status_name,"-",SME_lender_MIS$Partner.Sub.Status,"-",SME_lender_MIS$Reject.Remarks,"-",SME_lender_MIS$TSE.Comments)

SME_lender_MIS$Remark[SME_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',SME_lender_MIS$New_appops_status,ignore.case=TRUE) & SME_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, SME_lender_MIS$New_appops_status %in% c(280,380) 

SME_lender_MIS$Remark[SME_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',SME_lender_MIS$New_appops_status,ignore.case=TRUE) & SME_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, SME_lender_MIS$New_appops_status %in% c(480) 

SME_lender_MIS$Remark[SME_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',SME_lender_MIS$New_appops_status,ignore.case=TRUE) & SME_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, SME_lender_MIS$New_appops_status %in% c(580) 

SME_lender_MIS$Remark[SME_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',SME_lender_MIS$New_appops_status,ignore.case=TRUE) & SME_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, SME_lender_MIS$New_appops_status %in% c(680) 

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(SME_lender_MIS)[which(names(SME_lender_MIS) == "Approved.Amount")] <- "loan_amount"



SME_lender_MIS<-SME_lender_MIS %>% mutate('Lender'="Sme Corner",`Customer_requested_loan_amount`='')


write.xlsx(SME_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_SME_FB_File.xlsx")

SME_upload_1<-SME_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))





SME_upload_1<- SME_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=SME_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=SME_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(SME_upload_1, file = "C:\\R\\Lender Feedback\\Output\\SME_upload.xlsx")

