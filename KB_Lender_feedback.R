rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)
library(xtable)
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


KB_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\KB.xlsx", sheet = "CM_Base") %>% select(uId,State,first_loan_taken_date,first_loan_gmv,mobile,user_subState,Rejection_Reason,App_installed_date)

KB_lender_MIS_1<-KB_lender_MIS %>% filter(State %in% c('Confirmed'))

KB_lender_MIS_2<-KB_lender_MIS %>% filter(!State %in% c('Confirmed'))


#DIS_KB_lender_MIS
appos_map_KB <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'KREDITBEE')

KB_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Kredit Bee',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
KB_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Kredit Bee',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
KB_appos_dump<-rbind(KB_appos_dump1,KB_appos_dump2)
########################KB LenderFeedBack
KB_lender_MIS_1$mobile <- as.numeric(KB_lender_MIS_1$mobile)

#names(KB_lender_MIS_1)

colnames(KB_lender_MIS_1)[which(names(KB_lender_MIS_1) == "mobile")] <- "mobile_number"


colnames(KB_lender_MIS_1)[which(names(KB_lender_MIS_1) == "user_subState")] <- "CM_Status"

colnames(KB_lender_MIS_1)[which(names(KB_lender_MIS_1) == "first_loan_gmv")] <- "loan_amount"

KB_lender_MIS_1$offer_application_number<-KB_appos_dump$offer_application_number[match(KB_lender_MIS_1$mobile_number, KB_appos_dump$phone_home)]



KB_lender_MIS_1$CURRENT_appops_status<-KB_appos_dump$appops_status_code[match(KB_lender_MIS_1$mobile_number, KB_appos_dump$phone_home)]


KB_df_1<-left_join(KB_lender_MIS_1,appos_map_KB, by = c('State','CM_Status'))

colnames(KB_df_1)[which(names(KB_df_1) == "New_Status")] <- "New_appops_status"

#KB_lender_MIS_1$New_appops_status<-appos_map_KB$New_Status[match(KB_lender_MIS_1$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS_1$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS_1$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS_1$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS_1$New_appops_status, appos_map_KB$New_Status)]
#KB_df_1$latest_month <- format(as.Date(KB_df_1$first_loan_taken_date, format="%d/%m/%Y"),"%m")

#View(KB_lender_MIS_1)

current_mon<-format(as.Date(Sys.Date(), format="%d/%m/%Y"),"%m")
current_mon<-as.numeric(current_mon)

#last_mon<-as.numeric(current_mon)-1

#subset(KB_df_1, first_loan_taken_date > Sys.Date() - months(2))

# KB_df_1$first_loan_taken_date <- convertToDate(KB_df_1$first_loan_taken_date)
# 
# KB_df_1$first_loan_taken_date <- format(as.Date(KB_df_1$first_loan_taken_date), "%d-%m-%Y")

#View(KB_df_1)
#KB_df_1$first_loan_taken_date <- as.Date(KB_df_1$first_loan_taken_date, format = "%d/%m/%Y")

#class(KB_df_1$first_loan_taken_date)

KB_df_1 <-KB_df_1 %>%
  mutate(Remark = case_when(
    #grepl('Confirmed', State, ignore.case = T) & as.numeric(first_loan_taken_date == current_mon) ~ '990',# & grepl('Yes', first_loan_taken_or_not, ignore.case = T) ~ '990',
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

KB_df_1$Remark <- ifelse(is.na(KB_df_1$Remark), KB_df_1$New_appops_status, KB_df_1$Remark)

KB_df_1$Upload_Remarks <- str_c(KB_df_1$State,'-',KB_df_1$CM_Status,'-',KB_df_1$Rejection_Reason,'-',KB_df_1$NEW_appops_description)



KB_df_1$Remark[KB_df_1$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_1$New_appops_status,ignore.case=TRUE) & KB_df_1$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, KB_df_1$New_appops_status %in% c(280,380) 

KB_df_1$Remark[KB_df_1$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_1$New_appops_status,ignore.case=TRUE) & KB_df_1$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, KB_df_1$New_appops_status %in% c(480) 

KB_df_1$Remark[KB_df_1$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_1$New_appops_status,ignore.case=TRUE) & KB_df_1$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, KB_df_1$New_appops_status %in% c(580) 

KB_df_1$Remark[KB_df_1$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_1$New_appops_status,ignore.case=TRUE) & KB_df_1$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, KB_df_1$New_appops_status %in% c(680) 

colnames(KB_df_1)[which(names(KB_df_1) == "CM_Status")] <- "customer_status_name"

KB_df_1$Upload_Remarks <- ifelse(is.na(KB_df_1$Upload_Remarks), KB_df_1$customer_status_name, KB_df_1$Upload_Remarks)

#KB_df_1<-KB_df_1 %>% filter(!is.na(offer_application_number))

#############to chk
#KB_df_1<-KB_df_1 %>% filter(!is.na(offer_application_number) & !is.na(Remark))




colnames(KB_df_1)[which(names(KB_df_1) == "offer_application_number")] <- "Application_Number"

KB_df_1<-KB_df_1 %>% mutate('Lender'="KreditBee",`Customer_requested_loan_amount`='')


write.xlsx(KB_df_1, file = "C:\\R\\Lender Feedback\\Output\\Revised_KB_Disbursed_File.xlsx")

KB_upload_1<-KB_df_1 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

KB_upload_1<- KB_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=KB_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=KB_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=KB_upload_1$first_loan_gmv,`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



###############################


KB_lender_MIS_2$mobile <- as.numeric(KB_lender_MIS_2$mobile)

#names(KB_lender_MIS_2)

colnames(KB_lender_MIS_2)[which(names(KB_lender_MIS_2) == "mobile")] <- "mobile_number"


colnames(KB_lender_MIS_2)[which(names(KB_lender_MIS_2) == "user_subState")] <- "CM_Status"

colnames(KB_lender_MIS_2)[which(names(KB_lender_MIS_2) == "first_loan_gmv")] <- "loan_amount"

KB_lender_MIS_2$offer_application_number<-KB_appos_dump$offer_application_number[match(KB_lender_MIS_2$mobile_number, KB_appos_dump$phone_home)]



KB_lender_MIS_2$CURRENT_appops_status<-KB_appos_dump$appops_status_code[match(KB_lender_MIS_2$mobile_number, KB_appos_dump$phone_home)]


KB_df_2<-left_join(KB_lender_MIS_2,appos_map_KB, by = c('State','CM_Status'))

colnames(KB_df_2)[which(names(KB_df_2) == "New_Status")] <- "New_appops_status"

#KB_lender_MIS_2$New_appops_status<-appos_map_KB$New_Status[match(KB_lender_MIS_2$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS_2$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS_2$CM_Status, appos_map_KB$CM_Status)]

#KB_lender_MIS_2$NEW_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS_2$New_appops_status, appos_map_KB$New_Status)]
#KB_df_2$latest_month <- format(as.Date(KB_df_2$first_loan_taken_date, format="%d/%m/%Y"),"%m")

#View(KB_lender_MIS_2)

current_mon<-format(as.Date(Sys.Date(), format="%d/%m/%Y"),"%m")
current_mon<-as.numeric(current_mon)

#last_mon<-as.numeric(current_mon)-1

#subset(KB_df_2, first_loan_taken_date > Sys.Date() - months(2))

# KB_df_2$first_loan_taken_date <- convertToDate(KB_df_2$first_loan_taken_date)
# 
# KB_df_2$first_loan_taken_date <- format(as.Date(KB_df_2$first_loan_taken_date), "%d-%m-%Y")

#View(KB_df_2)
#KB_df_2$first_loan_taken_date <- as.Date(KB_df_2$first_loan_taken_date, format = "%d/%m/%Y")

#class(KB_df_2$first_loan_taken_date)

KB_df_2 <-KB_df_2 %>%
  mutate(Remark = case_when(
    #grepl('Confirmed', State, ignore.case = T) & as.numeric(first_loan_taken_date == current_mon) ~ '990',# & grepl('Yes', first_loan_taken_or_not, ignore.case = T) ~ '990',
    is.na(offer_application_number) ~ 'Not_in_CRM',
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

KB_df_2$Remark <- ifelse(is.na(KB_df_2$Remark), KB_df_2$New_appops_status, KB_df_2$Remark)

KB_df_2$Upload_Remarks <- str_c(KB_df_2$State,'-',KB_df_2$CM_Status,'-',KB_df_2$Rejection_Reason,'-',KB_df_2$NEW_appops_description)


KB_df_2$Remark[KB_df_2$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_2$New_appops_status,ignore.case=TRUE) & KB_df_2$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, KB_df_2$New_appops_status %in% c(280,380) 

KB_df_2$Remark[KB_df_2$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_2$New_appops_status,ignore.case=TRUE) & KB_df_2$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, KB_df_2$New_appops_status %in% c(480) 

KB_df_2$Remark[KB_df_2$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_2$New_appops_status,ignore.case=TRUE) & KB_df_2$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, KB_df_2$New_appops_status %in% c(580) 

KB_df_2$Remark[KB_df_2$Remark == 'Repeated_Feedback_cases' & grepl('80',KB_df_2$New_appops_status,ignore.case=TRUE) & KB_df_2$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, KB_df_2$New_appops_status %in% c(680) 

colnames(KB_df_2)[which(names(KB_df_2) == "CM_Status")] <- "customer_status_name"

KB_df_2$Upload_Remarks <- ifelse(is.na(KB_df_2$Upload_Remarks), KB_df_2$customer_status_name, KB_df_2$Upload_Remarks)

#KB_df_2<-KB_df_2 %>% filter(!is.na(offer_application_number))

#############to chk
#KB_df_2<-KB_df_2 %>% filter(!is.na(offer_application_number) & !is.na(Remark))




colnames(KB_df_2)[which(names(KB_df_2) == "offer_application_number")] <- "Application_Number"

KB_df_2<-KB_df_2 %>% mutate('Lender'="KreditBee", `Customer_requested_loan_amount`='')


write.xlsx(KB_df_2, file = "C:\\R\\Lender Feedback\\Output\\Revised_KB_Proposal_File.xlsx")

KB_upload_2<-KB_df_2 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

KB_upload_2<- KB_upload_2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=KB_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=KB_upload_2$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=KB_upload_2$first_loan_gmv,`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



#write.xlsx(KB_upload_1, file = "C:\\R\\Lender Feedback\\Output\\KB_upload_1.xlsx")





#write.xlsx(KB_upload_1, file = "C:\\R\\Lender Feedback\\Output\\KB_upload_1.xlsx")
