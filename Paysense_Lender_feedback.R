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

Paysense_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\Paysense.xlsx") %>% select(id,phone_id,user_status,application_current_status,amount)


appos_map_Paysense <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Paysense')

Paysense_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Paysense',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
Paysense_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Paysense',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
Paysense_appos_dump<-rbind(Paysense_appos_dump1,Paysense_appos_dump2)

########################Paysense LenderFeedBack

Paysense_lender_MIS$phone_id <- as.numeric(Paysense_lender_MIS$phone_id)

Paysense_lender_MIS$phone_id <- gsub('^91', '', Paysense_lender_MIS$phone_id)

colnames(Paysense_lender_MIS)[which(names(Paysense_lender_MIS) == "phone_id")] <- "mobile_number"

colnames(Paysense_lender_MIS)[which(names(Paysense_lender_MIS) == "application_current_status")] <- "customer_status_name"

colnames(Paysense_lender_MIS)[which(names(Paysense_lender_MIS) == "amount")] <- "loan_amount"


Paysense_lender_MIS$offer_application_number<-Paysense_appos_dump$offer_application_number[match(Paysense_lender_MIS$mobile_number, Paysense_appos_dump$phone_home)]



Paysense_lender_MIS$CURRENT_appops_status<-Paysense_appos_dump$appops_status_code[match(Paysense_lender_MIS$mobile_number, Paysense_appos_dump$phone_home)]


Paysense_lender_MIS$New_appops_status<-appos_map_Paysense$New_Status[match(Paysense_lender_MIS$customer_status_name, appos_map_Paysense$CM_Status)]

Paysense_lender_MIS$NEW_appops_description<-appos_map_Paysense$Status_Description[match(Paysense_lender_MIS$customer_status_name, appos_map_Paysense$CM_Status)]

names(Paysense_lender_MIS)
Paysense_lender_MIS <-Paysense_lender_MIS %>%
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

Paysense_lender_MIS$Remark <- ifelse(is.na(Paysense_lender_MIS$Remark), Paysense_lender_MIS$New_appops_status, Paysense_lender_MIS$Remark)

#Paysense_lender_MIS<-Paysense_lender_MIS %>% filter(!is.na(offer_application_number) & !is.na(Remark))
Paysense_lender_MIS$Upload_Remarks <- paste(Paysense_lender_MIS$user_status,'-',Paysense_lender_MIS$customer_status_name)


Paysense_lender_MIS$Remark[Paysense_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Paysense_lender_MIS$New_appops_status,ignore.case=TRUE) & Paysense_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, Paysense_lender_MIS$New_appops_status %in% c(280,380) 

Paysense_lender_MIS$Remark[Paysense_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Paysense_lender_MIS$New_appops_status,ignore.case=TRUE) & Paysense_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, Paysense_lender_MIS$New_appops_status %in% c(480) 

Paysense_lender_MIS$Remark[Paysense_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Paysense_lender_MIS$New_appops_status,ignore.case=TRUE) & Paysense_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, Paysense_lender_MIS$New_appops_status %in% c(580) 

Paysense_lender_MIS$Remark[Paysense_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Paysense_lender_MIS$New_appops_status,ignore.case=TRUE) & Paysense_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,690,700,701,780)] <- 680#, Paysense_lender_MIS$New_appops_status %in% c(680) 

colnames(Paysense_lender_MIS)[which(names(Paysense_lender_MIS) == "offer_application_number")] <- "Application_Number"


Paysense_lender_MIS<-Paysense_lender_MIS %>% mutate('Lender'="Paysense",`Customer_requested_loan_amount`='')

write.xlsx(Paysense_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Paysense_FB_File.xlsx")

Paysense_upload<-Paysense_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

Paysense_upload_1<-Paysense_upload %>% filter(Remark %in% c(990))

Paysense_upload_1<- Paysense_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Paysense_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Paysense_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`='75000',`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(Paysense_appos_dump, file = "C:\\R\\Lender Feedback\\Output\\Paysense_appos_dump.xlsx")


Paysense_upload_2<-Paysense_upload %>% filter(!Remark %in% c(990))

Paysense_upload_2<- Paysense_upload_2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Paysense_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Paysense_upload_2$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(Paysense_upload_2, file = "C:\\R\\Lender Feedback\\Output\\Paysense_upload_2.xlsx")

#################################
