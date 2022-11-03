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
library(logging)
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
#SHOULD SAVE THE FILEIN XLSX

Faircent_lender_MIS<- read.xlsx("C:\\R\\Lender Feedback\\Input\\Faircent.xlsx") %>% select(Contact.Number,Loan.Status,Loan.Amount,Disbursement.Date,User.Call.Status,Created.Date)




#names(Faircent_lender_MIS)
appos_map_Faircent <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Faircent')
Faircent_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Faircent',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
Faircent_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Faircent',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)

Faircent_appos_dump<-rbind(Faircent_appos_dump1,Faircent_appos_dump2)
########################Faircent LenderFeedBack

Faircent_lender_MIS$Contact.Number <- as.numeric(Faircent_lender_MIS$Contact.Number)

colnames(Faircent_lender_MIS)[which(names(Faircent_lender_MIS) == "Contact.Number")] <- "mobile_number"

colnames(Faircent_lender_MIS)[which(names(Faircent_lender_MIS) == "Loan.Status")] <- "customer_status_name"

colnames(Faircent_lender_MIS)[which(names(Faircent_lender_MIS) == "Loan.Amount")] <- "loan_amount"

#Faircent_appos_dump$ph_no<-Faircent_lender_MIS$mobile_number[match(Faircent_appos_dump$phone_home, Faircent_lender_MIS$mobile_number)]

Faircent_lender_MIS$offer_application_number<-Faircent_appos_dump$offer_application_number[match(Faircent_lender_MIS$mobile_number, Faircent_appos_dump$phone_home)]



Faircent_lender_MIS$CURRENT_appops_status<-Faircent_appos_dump$appops_status_code[match(Faircent_lender_MIS$mobile_number, Faircent_appos_dump$phone_home)]

Faircent_lender_MIS$New_appops_status<-appos_map_Faircent$New_Status[match(Faircent_lender_MIS$customer_status_name, appos_map_Faircent$CM_Status)]

Faircent_lender_MIS$NEW_appops_description<-appos_map_Faircent$Status_Description[match(Faircent_lender_MIS$customer_status_name, appos_map_Faircent$CM_Status)]


Faircent_lender_MIS <-Faircent_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                      300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,690,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

Faircent_lender_MIS$Remark <- ifelse(is.na(Faircent_lender_MIS$Remark), Faircent_lender_MIS$New_appops_status, Faircent_lender_MIS$Remark)

Faircent_lender_MIS$Upload_Remarks <- paste(Faircent_lender_MIS$customer_status_name,"-",Faircent_lender_MIS$User.Call.Status)

Faircent_lender_MIS$Remark[Faircent_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Faircent_lender_MIS$New_appops_status,ignore.case=TRUE) & Faircent_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                     300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, Faircent_lender_MIS$New_appops_status %in% c(280,380) 

Faircent_lender_MIS$Remark[Faircent_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Faircent_lender_MIS$New_appops_status,ignore.case=TRUE) & Faircent_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, Faircent_lender_MIS$New_appops_status %in% c(480) 

Faircent_lender_MIS$Remark[Faircent_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Faircent_lender_MIS$New_appops_status,ignore.case=TRUE) & Faircent_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, Faircent_lender_MIS$New_appops_status %in% c(580) 

Faircent_lender_MIS$Remark[Faircent_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Faircent_lender_MIS$New_appops_status,ignore.case=TRUE) & Faircent_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, Faircent_lender_MIS$New_appops_status %in% c(680) 

colnames(Faircent_lender_MIS)[which(names(Faircent_lender_MIS) == "offer_application_number")] <- "Application_Number"
colnames(Faircent_lender_MIS)[which(names(Faircent_lender_MIS) == "Loan.Amount")] <- "loan_amount"

Faircent_lender_MIS<-Faircent_lender_MIS %>% mutate('Lender'="Faircent",`Customer_requested_loan_amount`='')


write.xlsx(Faircent_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Faircent_FB_File.xlsx")

Faircent_upload_1<-Faircent_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))





Faircent_upload_1<- Faircent_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Faircent_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Faircent_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(Faircent_upload_1, file = "C:\\R\\Lender Feedback\\Output\\Faircent_upload.xlsx")

