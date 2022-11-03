rm(list=ls())
#library(magrittr)
library(tibble)
library(dplyr)
library(plyr)
library(tidyr)
library(lubridate)
library(openxlsx)
library(tidyverse)
library(data.table)

library(purrr)
library(janitor)


#AXIS1 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Axis_Disbursed_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

#AXIS2 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Axis_Proposal_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

CASHE <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_CashE_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

CITI <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_CITI_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

ES <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_ES_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

FAIR <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Faircent_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

FLEXI <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_FLEXI_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

HC <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_HC_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)


#IDFC <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_IDFC_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

#IIFL <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_IIFL_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

INDIF <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_INDIFI_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

KB1 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_KB_Disbursed_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

KB2 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_KB_Proposal_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

LK <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_LK_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

#LNT <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_LNT_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MCAP <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_MCAP_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MUTH <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Muthoot_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MV1 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_MV_Disbursed_FB_File_1.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MV2 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_MV_Disbursed_FB_File_2.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MV3 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_MV_Proposal_FB_File_1.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

MV4 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_MV_Proposal_FB_File_2.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)


PAY <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Paysense_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

PRO <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_Protium_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

#RBL <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_RBL_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

SBI <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_SBI_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

SCUF1 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_SCUF_Disbursed_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

SCUF2 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_SCUF_Proposal_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

SCUF3 <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_SCUF_na.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

YES <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_YES_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

SME <- read.xlsx("C:\\R\\Lender Feedback\\Output\\2022-11-01\\Revised_SME_FB_File.xlsx") %>% select(Lender,mobile_number,Application_Number,loan_amount,Customer_requested_loan_amount,CURRENT_appops_status,New_appops_status,NEW_appops_description,Remark,Upload_Remarks)

merge_df<-rbind(CASHE,CITI,ES,FAIR,FLEXI,HC,INDIF,KB1,KB2,LK,MCAP,MUTH,MV1,MV2,MV3,MV4,PRO,PAY,SBI,SCUF1,SCUF2,SCUF3,SME,YES)#CITI,KB1,KB2,PROTI,AXIS1,AXIS2,SBI,YES,IIFL,MV1,MV2,MV3,MV4

#names(desc_merge_df)

desc_merge_df <- merge_df %>% group_by(Lender,Remark) %>% dplyr::summarise(Total = n_distinct(mobile_number, na.rm = TRUE))# %>% adorn_totals("row")

desc_merge_df<-spread(desc_merge_df,key=Remark, value = Total)

desc_merge_df<-desc_merge_df %>% adorn_totals("row")

write.xlsx(desc_merge_df, file = "C:\\R\\Lender Feedback\\Output\\Final_FB_MIS.xlsx")

write.xlsx(merge_df, file = "C:\\R\\Lender Feedback\\Output\\Final_FB_MIS_dump.xlsx")





#########################
thisdate<-format(Sys.Date(),'%Y-%m-%d')

not_in_crm<-merge_df %>% filter(Remark %in% c('Not_in_CRM'), is.na(Application_Number))

merge_df<-merge_df %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM','Status_not_received','Stuck_cases'), !is.na(Application_Number))

FB_upload_1<-merge_df %>% filter(!Remark %in% c('690','990'))

FB_upload_1<- FB_upload_1 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=FB_upload_1$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=FB_upload_1$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`='',`Booking_Date`='',`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(FB_upload_1, file = "C:\\R\\Lender Feedback\\Output\\FB_upload_1.xlsx")


FB_upload_2<-merge_df %>% filter(Remark %in% c('690'))

FB_upload_2<- FB_upload_2 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=FB_upload_2$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=FB_upload_2$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`=FB_upload_2$Customer_requested_loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(FB_upload_2, file = "C:\\R\\Lender Feedback\\Output\\FB_upload_2.xlsx")


FB_upload_3<-merge_df %>% filter(Remark %in% c('990'))

FB_upload_3<- FB_upload_3 %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=FB_upload_3$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=FB_upload_3$Upload_Remarks,`Offer_Reference_Number`='',`Loan_Sanctioned_Disbursed_Amount`=FB_upload_3$loan_amount,Customer_requested_loan_amount,`Booking_Date`=thisdate,`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(FB_upload_3, file = "C:\\R\\Lender Feedback\\Output\\FB_upload_3.xlsx")


#################################

FB_df<-not_in_crm %>% filter(Remark %in% c('Not_in_CRM'))

write.csv(FB_df, file = "C:\\R\\Lender Feedback\\Output\\Not_in_CRM.csv")

NOT_IN <- read.csv("C:\\R\\Lender Feedback\\Output\\application_no-2022-10-31.csv")

NOT_IN$phone_home <- as.numeric(NOT_IN$phone_home)

NOT_IN <-NOT_IN %>%
  mutate(Lender = case_when(
    grepl('CASHE|CashE', lender, ignore.case = T) ~ 'CashE',
    grepl('Early Salary|EARLYSALARYPL', lender, ignore.case = T) ~ 'Early Salary',
    grepl('Lending Kart', lender, ignore.case = T) ~ 'Lendingkart',
    grepl('MONEYVIEW', lender, ignore.case = T) ~ 'MoneyView',
    grepl('MUTHOOT', lender, ignore.case = T) ~ 'MUTHOOT'
    
  ))


colnames(NOT_IN)[which(names(NOT_IN) == "phone_home")] <- "mobile_number"

NOT_IN$mobile_number <- as.numeric(NOT_IN$mobile_number)

FB_df$mobile_number <- as.numeric(FB_df$mobile_number)

FINAL_df<-left_join(FB_df,NOT_IN, by = c('Lender','mobile_number'))

write.xlsx(FINAL_df, file = "C:\\R\\Lender Feedback\\Output\\FINAL_df.xlsx")
