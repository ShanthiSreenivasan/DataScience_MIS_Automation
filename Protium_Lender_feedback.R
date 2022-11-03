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


#NEED to del dup col appl1_mobile in the FB xls
Protium_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Protium.xlsx", sheet='ApplicationData') %>% select(appl1_mobile,stage,substage,model_amount)

#names(Protium_lender_MIS)
appos_map_Protium <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Protium')


Protium_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Protium',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
Protium_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Protium',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
Protium_appos_dump<-rbind(Protium_appos_dump1,Protium_appos_dump2)

########################Protium LenderFeedBack

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "appl1_mobile")] <- "mobile_number"

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "stage")] <- "customer_status_name"

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "model_amount")] <- "model_amount"


Protium_lender_MIS$mobile_number <- as.numeric(Protium_lender_MIS$mobile_number)


Protium_lender_MIS$offer_application_number<-Protium_appos_dump$offer_application_number[match(Protium_lender_MIS$mobile_number, Protium_appos_dump$phone_home)]



Protium_lender_MIS$CURRENT_appops_status<-Protium_appos_dump$appops_status_code[match(Protium_lender_MIS$mobile_number, Protium_appos_dump$phone_home)]


Protium_lender_MIS$New_appops_status<-appos_map_Protium$New_Status[match(Protium_lender_MIS$customer_status_name, appos_map_Protium$CM_Status)]

Protium_lender_MIS$NEW_appops_description<-appos_map_Protium$Status_Description[match(Protium_lender_MIS$customer_status_name, appos_map_Protium$CM_Status)]


Protium_lender_MIS <-Protium_lender_MIS %>%
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
    is.na(New_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

Protium_lender_MIS$Remark <- ifelse(is.na(Protium_lender_MIS$Remark), Protium_lender_MIS$New_appops_status, Protium_lender_MIS$Remark)

Protium_lender_MIS$Upload_Remarks <- paste(Protium_lender_MIS$customer_status_name,'-',Protium_lender_MIS$substage)


Protium_lender_MIS$Remark[Protium_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Protium_lender_MIS$New_appops_status,ignore.case=TRUE) & Protium_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                       300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, Protium_lender_MIS$New_appops_status %in% c(280,380) 

Protium_lender_MIS$Remark[Protium_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Protium_lender_MIS$New_appops_status,ignore.case=TRUE) & Protium_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, Protium_lender_MIS$New_appops_status %in% c(480) 

Protium_lender_MIS$Remark[Protium_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Protium_lender_MIS$New_appops_status,ignore.case=TRUE) & Protium_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, Protium_lender_MIS$New_appops_status %in% c(580) 

Protium_lender_MIS$Remark[Protium_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',Protium_lender_MIS$New_appops_status,ignore.case=TRUE) & Protium_lender_MIS$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, Protium_lender_MIS$New_appops_status %in% c(680) 

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "offer_application_number")] <- "Application_Number"

colnames(Protium_lender_MIS)[which(names(Protium_lender_MIS) == "model_amount")] <- "loan_amount"

Protium_lender_MIS<-Protium_lender_MIS %>% mutate('Lender'="Protium",`Customer_requested_loan_amount`='')

write.xlsx(Protium_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Protium_FB_File.xlsx")

Protium_upload<-Protium_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

Protium_upload<- Protium_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Protium_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Protium_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(Protium_upload, file = "C:\\R\\Lender Feedback\\Output\\Protium_upload.xlsx")




#################################
