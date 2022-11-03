rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)
library(xtable)
library(data.table)

library(stringr)
library(lubridate)
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
  dir.create(paste("./Output/Revised_FB_File/",thisdate,sep=""))
} 

DIS_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV1.xlsx", sheet = "disbursals") %>% select(mobile_number,loan_amount,loan_amount)


#PRO_MV_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV.xlsx", sheet = "Sheet1")

#DIS_MV_lender_MIS
appos_map_MV <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx", sheet = 'MoneyView')


MV_appos_dump1 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_1.csv") %>% filter(grepl('Money View',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)
MV_appos_dump2 <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump_2.csv") %>% filter(grepl('Money View',name,ignore.case=TRUE)) %>% select(lead_id,phone_home,offer_application_number,status,name,appops_status_code)

MV_appos_dump<-rbind(MV_appos_dump1,MV_appos_dump2)

########################MV LenderFeedBack

DIS_MV_lender_MIS$loan_amount <- as.numeric(DIS_MV_lender_MIS$loan_amount)

#colnames(DIS_MV_lender_MIS)[which(names(DIS_MV_lender_MIS) == "PERMOBILE")] <- "mobile_number"


MV_appos_dump$ph_no<-DIS_MV_lender_MIS$mobile_number[match(MV_appos_dump$phone_home, DIS_MV_lender_MIS$mobile_number)]

DIS_MV_lender_MIS$offer_application_number<-MV_appos_dump$offer_application_number[match(DIS_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]



DIS_MV_lender_MIS$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(DIS_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]


DIS_MV_lender_MIS$New_appops_status<- '990'

DIS_MV_lender_MIS$NEW_appops_description<-"Loan Disbursed"

#DIS_MV_lender_MIS$NEW_appops_description<-appos_map_MV$Description[match(DIS_MV_lender_MIS$customer_status_name, appos_map_MV$Campaign_Sub_Status)]

DIS_MV_lender_MIS <-DIS_MV_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(offer_application_number) ~ 'Not_in_CRM'
    
    
  ))

DIS_MV_lender_MIS$Remark <- ifelse(is.na(DIS_MV_lender_MIS$Remark), DIS_MV_lender_MIS$New_appops_status, DIS_MV_lender_MIS$Remark)


DIS_MV_lender_MIS$Upload_Remarks <- str_c(DIS_MV_lender_MIS$NEW_appops_description,' ',DIS_MV_lender_MIS$loan_amount)


colnames(DIS_MV_lender_MIS)[which(names(DIS_MV_lender_MIS) == "offer_application_number")] <- "Application_Number"

DIS_MV_lender_MIS<-DIS_MV_lender_MIS %>% mutate('Lender'="MoneyView",`Customer_requested_loan_amount`='')

write.xlsx(DIS_MV_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Disbursed_FB_File_1.xlsx")

MV_upload_1<-DIS_MV_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))


MV_upload_1<- MV_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=MV_upload_1$loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



#write.xlsx(MV_upload_1, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_1_1.xlsx")


######################


Status_MV_lender_MIS_1 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV1.xlsx", sheet = "status") %>% select(Mobile,Disbursed_amt,Offer_selected,Rejected,Cancelled,Submits,In_Process,NACH,Disbursal_initiated,Loan_Approved)


colnames(Status_MV_lender_MIS_1)[which(names(Status_MV_lender_MIS_1) == "Disbursed_amt")] <- "loan_amount"

offer<-Status_MV_lender_MIS_1 %>% filter(Offer_selected==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Offer_selected")


Rejected<-Status_MV_lender_MIS_1 %>% filter(Rejected==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Rejected")

Cancelled<-Status_MV_lender_MIS_1 %>% filter(Cancelled==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Cancelled")

Submits<-Status_MV_lender_MIS_1 %>% filter(Submits==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Submits")

In_Process<-Status_MV_lender_MIS_1 %>% filter(In_Process==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="In_Process")

NACH<-Status_MV_lender_MIS_1 %>% filter(NACH==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="NACH")

Loan_Approved<-Status_MV_lender_MIS_1 %>% filter(Loan_Approved==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Loan_Approved")

Disbursal_initiated<-Status_MV_lender_MIS_1 %>% filter(Disbursal_initiated==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Disbursal_initiated")


PRO_MV_lender_MIS<-rbind(offer,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved,Disbursal_initiated)

colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "Mobile")] <- "mobile_number"


colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "REMARKS")] <- "customer_status_name"


#PRO_MV_lender_MIS$mobile_number <- as.numeric(PRO_MV_lender_MIS$mobile_number)


PRO_MV_lender_MIS$ph_no<-DIS_MV_lender_MIS$mobile_number[match(PRO_MV_lender_MIS$mobile_number, DIS_MV_lender_MIS$mobile_number)]
PRO_MV_lender_MIS$ph_no[is.na(PRO_MV_lender_MIS$ph_no)] <- 'FALSE'

PRO_MV_lender_MIS<-PRO_MV_lender_MIS %>% filter(ph_no=='FALSE')

MV_appos_dump$ph_no<-PRO_MV_lender_MIS$mobile_number[match(MV_appos_dump$phone_home, PRO_MV_lender_MIS$mobile_number)]

PRO_MV_lender_MIS$offer_application_number<-MV_appos_dump$offer_application_number[match(PRO_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]



PRO_MV_lender_MIS$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(PRO_MV_lender_MIS$mobile_number, MV_appos_dump$phone_home)]

PRO_MV_lender_MIS$New_appops_status<-appos_map_MV$New_Status[match(PRO_MV_lender_MIS$customer_status_name, appos_map_MV$CM_Status)]

PRO_MV_lender_MIS$NEW_appops_description<-appos_map_MV$Status_Description[match(PRO_MV_lender_MIS$customer_status_name, appos_map_MV$CM_Status)]


PRO_MV_lender_MIS <-PRO_MV_lender_MIS %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    (New_appops_status %in% c(280,380) & CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,300,320,350,360,370,382,383,390,393,394,395,399)) ~ '380',
    (New_appops_status %in% c(480) & CURRENT_appops_status %in% c(400,401,420,421,450,460,470,490,491,494,495)) ~ '480',
    (New_appops_status %in% c(580) & CURRENT_appops_status %in% c(500,520,550,560,570,590)) ~ '580',
    (New_appops_status %in% c(680) & CURRENT_appops_status %in% c(650,660,670,700,701,780)) ~ '680',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(New_appops_status) ~'Status_not_received',
    is.na(CURRENT_appops_status) & is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))


PRO_MV_lender_MIS$Remark <- ifelse(is.na(PRO_MV_lender_MIS$Remark), PRO_MV_lender_MIS$New_appops_status, PRO_MV_lender_MIS$Remark)



PRO_MV_lender_MIS$Upload_Remarks <- str_c(PRO_MV_lender_MIS$customer_status_name,'-',PRO_MV_lender_MIS$NEW_appops_description)


PRO_MV_lender_MIS$Remark[PRO_MV_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, PRO_MV_lender_MIS$New_appops_status %in% c(280,380) 

PRO_MV_lender_MIS$Remark[PRO_MV_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, PRO_MV_lender_MIS$New_appops_status %in% c(480) 

PRO_MV_lender_MIS$Remark[PRO_MV_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, PRO_MV_lender_MIS$New_appops_status %in% c(580) 

PRO_MV_lender_MIS$Remark[PRO_MV_lender_MIS$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS$CURRENT_appops_status %in% c(650,660,670,690,700,701,780)] <- 680#, PRO_MV_lender_MIS$New_appops_status %in% c(680) 

colnames(PRO_MV_lender_MIS)[which(names(PRO_MV_lender_MIS) == "offer_application_number")] <- "Application_Number"


PRO_MV_lender_MIS<-PRO_MV_lender_MIS %>% mutate('Lender'="MoneyView",`Customer_requested_loan_amount`='')

write.xlsx(PRO_MV_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Proposal_FB_File_1.xlsx") 

MV_upload_2<-PRO_MV_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

MV_upload_2<- MV_upload_2 %>% filter(!grepl('Repeated_Feedback_cases, Not_in_CRM',Remark,ignore.case=TRUE)) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(MV_upload_2, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_2_1.xlsx")
##########################################################################



DIS_MV_lender_MIS_2 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV2.xlsx", sheet = "disbursals") %>% select(mobile_number,loan_amount,loan_amount)

########################MV LenderFeedBack

DIS_MV_lender_MIS_2$loan_amount <- as.numeric(DIS_MV_lender_MIS_2$loan_amount)

#colnames(DIS_MV_lender_MIS_2)[which(names(DIS_MV_lender_MIS_2) == "PERMOBILE")] <- "mobile_number"


MV_appos_dump$ph_no<-DIS_MV_lender_MIS_2$mobile_number[match(MV_appos_dump$phone_home, DIS_MV_lender_MIS_2$mobile_number)]

DIS_MV_lender_MIS_2$offer_application_number<-MV_appos_dump$offer_application_number[match(DIS_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]



DIS_MV_lender_MIS_2$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(DIS_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]


DIS_MV_lender_MIS_2$New_appops_status<- '990'

DIS_MV_lender_MIS_2$NEW_appops_description<-"Loan Disbursed"

#DIS_MV_lender_MIS_2$NEW_appops_description<-appos_map_MV$Description[match(DIS_MV_lender_MIS_2$customer_status_name, appos_map_MV$Campaign_Sub_Status)]

DIS_MV_lender_MIS_2 <-DIS_MV_lender_MIS_2 %>%
  mutate(Remark = case_when(
    CURRENT_appops_status %in% c(710) ~ 'Stuck_cases',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    (New_appops_status ==990) ~ '990',
    is.na(offer_application_number) ~ 'Not_in_CRM'
    
  ))

DIS_MV_lender_MIS_2$Remark <- ifelse(is.na(DIS_MV_lender_MIS_2$Remark), DIS_MV_lender_MIS_2$New_appops_status, DIS_MV_lender_MIS_2$Remark)


DIS_MV_lender_MIS_2$Upload_Remarks <- str_c(DIS_MV_lender_MIS_2$NEW_appops_description,' ',DIS_MV_lender_MIS_2$loan_amount)


colnames(DIS_MV_lender_MIS_2)[which(names(DIS_MV_lender_MIS_2) == "offer_application_number")] <- "Application_Number"

DIS_MV_lender_MIS_2<-DIS_MV_lender_MIS_2 %>% mutate('Lender'="MoneyView",`Customer_requested_loan_amount`='')

write.xlsx(DIS_MV_lender_MIS_2, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Disbursed_FB_File_2.xlsx")

MV_upload_1<-DIS_MV_lender_MIS_2 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))


MV_upload_1<- MV_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_1$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`=MV_upload_1$loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")



#write.xlsx(MV_upload_1, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_1_2.xlsx")


######################




Status_MV_lender_MIS_2 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\MV2.xlsx", sheet = "status") %>% select(Mobile,Disbursed_amt,Offer_selected,Rejected,Cancelled,Submits,In_Process,NACH,Disbursal_initiated,Loan_Approved)


colnames(Status_MV_lender_MIS_2)[which(names(Status_MV_lender_MIS_2) == "Disbursed_amt")] <- "loan_amount"

offer<-Status_MV_lender_MIS_2 %>% filter(Offer_selected==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Offer_selected")


Rejected<-Status_MV_lender_MIS_2 %>% filter(Rejected==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Rejected")

Cancelled<-Status_MV_lender_MIS_2 %>% filter(Cancelled==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Cancelled")

Submits<-Status_MV_lender_MIS_2 %>% filter(Submits==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Submits")

In_Process<-Status_MV_lender_MIS_2 %>% filter(In_Process==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="In_Process")

NACH<-Status_MV_lender_MIS_2 %>% filter(NACH==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="NACH")

Loan_Approved<-Status_MV_lender_MIS_2 %>% filter(Loan_Approved==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Loan_Approved")

Disbursal_initiated<-Status_MV_lender_MIS_2 %>% filter(Disbursal_initiated==1) %>% select(Mobile,loan_amount) %>% mutate(`REMARKS`="Disbursal_initiated")


PRO_MV_lender_MIS_2<-rbind(offer,Rejected,Cancelled,Submits,In_Process,NACH,Loan_Approved,Disbursal_initiated)

colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "Mobile")] <- "mobile_number"


colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "REMARKS")] <- "customer_status_name"


#PRO_MV_lender_MIS_2$mobile_number <- as.numeric(PRO_MV_lender_MIS_2$mobile_number)


PRO_MV_lender_MIS_2$ph_no<-DIS_MV_lender_MIS_2$mobile_number[match(PRO_MV_lender_MIS_2$mobile_number, DIS_MV_lender_MIS_2$mobile_number)]
PRO_MV_lender_MIS_2$ph_no[is.na(PRO_MV_lender_MIS_2$ph_no)] <- 'FALSE'

PRO_MV_lender_MIS_2<-PRO_MV_lender_MIS_2 %>% filter(ph_no=='FALSE')

MV_appos_dump$ph_no<-PRO_MV_lender_MIS_2$mobile_number[match(MV_appos_dump$phone_home, PRO_MV_lender_MIS_2$mobile_number)]

PRO_MV_lender_MIS_2$offer_application_number<-MV_appos_dump$offer_application_number[match(PRO_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]



PRO_MV_lender_MIS_2$CURRENT_appops_status<-MV_appos_dump$appops_status_code[match(PRO_MV_lender_MIS_2$mobile_number, MV_appos_dump$phone_home)]

PRO_MV_lender_MIS_2$New_appops_status<-appos_map_MV$New_Status[match(PRO_MV_lender_MIS_2$customer_status_name, appos_map_MV$CM_Status)]

PRO_MV_lender_MIS_2$NEW_appops_description<-appos_map_MV$Status_Description[match(PRO_MV_lender_MIS_2$customer_status_name, appos_map_MV$CM_Status)]


PRO_MV_lender_MIS_2 <-PRO_MV_lender_MIS_2 %>%
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

PRO_MV_lender_MIS_2$Remark <- ifelse(is.na(PRO_MV_lender_MIS_2$Remark), PRO_MV_lender_MIS_2$New_appops_status, PRO_MV_lender_MIS_2$Remark)



PRO_MV_lender_MIS_2$Upload_Remarks <- str_c(PRO_MV_lender_MIS_2$customer_status_name,'-',PRO_MV_lender_MIS_2$NEW_appops_description)


PRO_MV_lender_MIS_2$Remark[PRO_MV_lender_MIS_2$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS_2$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS_2$CURRENT_appops_status %in% c(89,140,150,210,212,213,250,270,272,
                                                                                                                                                                                                           300,320,350,360,370,382,383,390,393,394,395,399)] <- 380 #, PRO_MV_lender_MIS_2$New_appops_status %in% c(280,380) 

PRO_MV_lender_MIS_2$Remark[PRO_MV_lender_MIS_2$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS_2$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS_2$CURRENT_appops_status %in% c(400,401,420,421,425,450,460,470,490,491,494,495)] <- 480 #, PRO_MV_lender_MIS_2$New_appops_status %in% c(480) 

PRO_MV_lender_MIS_2$Remark[PRO_MV_lender_MIS_2$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS_2$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS_2$CURRENT_appops_status %in% c(500,520,550,560,570,590)] <- 580#, PRO_MV_lender_MIS_2$New_appops_status %in% c(580) 

PRO_MV_lender_MIS_2$Remark[PRO_MV_lender_MIS_2$Remark == 'Repeated_Feedback_cases' & grepl('80',PRO_MV_lender_MIS_2$New_appops_status,ignore.case=TRUE) & PRO_MV_lender_MIS_2$CURRENT_appops_status %in% c(690,650,660,670,700,701,780)] <- 680#, PRO_MV_lender_MIS_2$New_appops_status %in% c(680) 

colnames(PRO_MV_lender_MIS_2)[which(names(PRO_MV_lender_MIS_2) == "offer_application_number")] <- "Application_Number"


PRO_MV_lender_MIS_2<-PRO_MV_lender_MIS_2 %>% mutate('Lender'="MoneyView",`Customer_requested_loan_amount`='')

write.xlsx(PRO_MV_lender_MIS_2, file = "C:\\R\\Lender Feedback\\Output\\Revised_MV_Proposal_FB_File_2.xlsx") 

MV_upload_2<-PRO_MV_lender_MIS_2 %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Stuck_cases'), !is.na(Application_Number))

MV_upload_2<- MV_upload_2 %>% filter(!grepl('Repeated_Feedback_cases, Not_in_CRM',Remark,ignore.case=TRUE)) %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=MV_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=MV_upload_2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


#write.xlsx(MV_upload_2, file = "C:\\R\\Lender Feedback\\Output\\MV_upload_2_2.xlsx")


