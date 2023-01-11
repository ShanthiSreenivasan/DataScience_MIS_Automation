rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)

library(stringr)
library(lubridate)
library(openxlsx)
library(purrr)
library(janitor)

options(scipen = 999)

setwd("C:\\R\\Collection\\Input")
# 
# getwd()
# 
# thisdate<-format(Sys.Date(),'%Y-%m-%d')
# 
# source('.\\Function file.R')
# 
# if(!dir.exists(paste("./Output/",thisdate,sep="")))
# {
#   dir.create(paste("./Output/",thisdate,sep=""))
# } 

RP_dump <- read.xlsx("C:\\R\\Collection\\Input\\R&P.xlsx", sheet= "Sheet1") %>% select(id,lead_id,lender,product,account_status,amount_paid,payment_date,Customer_name,Account_No,AAN_no)

RP_dump$lead_id <- as.numeric(RP_dump$lead_id)


WH_dump <- read.csv("C:\\R\\Collection\\Input\\CIS Dump_WH.csv") %>% select(lead_id,lender,product,account_status,Account_No,Alternate_acc_no)

#WH_dump$Account_No<-gsub("`","",Collection_df$Account_No, ignore.case = TRUE)

#WH_dump$Alternate_acc_no<-gsub("`","",Collection_df$Alternate_acc_no, ignore.case = TRUE)



Collection_dump <- read.csv("C:\\R\\Collection\\Input\\COLLECTION OPS PAYMENT FILE.csv") %>% select(id,lead_id,lender,product,account_status,amount_paid,payment_date,first_name,last_name)


# df1 <- read.csv("data.csv", colClasses=c("numeric", "Date", "character"))
# df2 <- read.csv("data.csv", colClasses=c("numeric", "Date", "character"))


Collection_dump$Customer_name <- str_c(Collection_dump$first_name,' ',Collection_dump$last_name)

Collection_df<-Collection_dump%>% select(id,lead_id,lender,product,account_status,amount_paid,payment_date,Customer_name)

Collection_df$Account_No<-WH_dump$Account_No[match(Collection_dump$lead_id, WH_dump$lead_id)]

Collection_df$AAN_no<-WH_dump$Alternate_acc_no[match(Collection_dump$lead_id, WH_dump$lead_id)]

#Collection_df$Account_No<-gsub("`",'',Collection_df$Account_No, ignore.case = TRUE)
Collection_df$Account_No<-str_replace(Collection_df$Account_No, '` ',"")
#Collection_df$AAN_no<-gsub("`",'',Collection_df$AAN_no, ignore.case = TRUE)
Collection_df$AAN_no<-str_replace(Collection_df$AAN_no, '` ',"")

Collection_df[!as.numeric(Collection_df$lead_id) %in% as.numeric(RP_dump$lead_id), , drop = FALSE]

Final_Collection<-bind_rows(Collection_df,RP_dump)

Final_Collection<-Final_Collection[!duplicated(Final_Collection$id), ]

AAN_dump <- read.xlsx("C:\\R\\Collection\\Input\\AAN Master.xlsx", sheet= "Consolidated")



IP_file <- read.xlsx("C:\\R\\Collection\\Input\\DPR_Input_file.xlsx")

IP_file$AAN_no <- as.numeric(IP_file$AAN_no)

#IP_file$Account_No<-str_replace(IP_file$Account_No, '` ',"")

#class(Final_Collection$AAN_no)
Final_Collection$dup_chk<-IP_file$Account_No[match(as.numeric(Final_Collection$Account_No), as.numeric(IP_file$Account_No))]

#Final_Collection$dup_chk<-IP_file$Account_No[match(Final_Collection$Account_No, IP_file$Account_No)]

#names(AAN_dump)

Final_Collection$dup_chk2<-IP_file$AAN_no[match(as.numeric(Final_Collection$AAN_no), as.numeric(IP_file$AAN_no))]

#Final_Collection$dup_chk2<-IP_file$AAN_no[match(Final_Collection$AAN_no, IP_file$AAN_no)]

Final_Collection$AAN_no2<-AAN_dump$AAN[match(as.numeric(Final_Collection$Account_No), as.numeric(AAN_dump$Account.number))]

Final_Collection$dup_chk3<-IP_file$AAN_no[match(as.numeric(Final_Collection$AAN_no2), as.numeric(IP_file$AAN_no))]



Final_Collection <-Final_Collection %>%
  mutate(dup_cases_flag = case_when(
    is.na(dup_chk) ~ '0',
    !is.na(dup_chk) ~ '1'
  ))

Final_Collection <-Final_Collection %>%
  mutate(dup_cases_flag2 = case_when(
    is.na(dup_chk2) ~ '0',
    !is.na(dup_chk2) ~ '1'
  ))
Final_Collection <-Final_Collection %>% 
  mutate(dup_cases_flag3 = case_when(
    is.na(dup_chk3) ~ '0',
    !is.na(dup_chk3) ~ '1'
  ))



Final_Collection<- Final_Collection %>% filter(dup_cases_flag == 0, dup_cases_flag2 == 0, dup_cases_flag3 == 0)



Final_Collection<-Final_Collection %>% select(id,lead_id,lender,product,account_status,amount_paid,payment_date,Customer_name,Account_No,AAN_no)
#names(Final_Collection)
#Final_Collection$Account_No<-sub("^","`",Final_Collection$Account_No)

#Final_Collection$AAN_no<-sub("^","`",Final_Collection$AAN_no)


# Final_Collection <- Final_Collection %>%
#   mutate(lender = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & grepl('CC|CCC|credit card|CL', product, ignore.case = T) ~ 'HDFC CC', 
#                             TRUE ~ as.character(lender)),
#          lender = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & !grepl('CC|CCC|credit card|CL', product, ignore.case = T) ~ 'HDFC Retail', 
#                             TRUE ~ as.character(lender)), 
#          lender = lender)

Final_Collection <- Final_Collection %>%
  mutate(hdfc_flag = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & grepl('CC|CCC|credit card|CL', product, ignore.case = T) ~ 'HDFC CC', 
                            TRUE ~ as.character(lender)),
         hdfc_flag = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & !grepl('CC|CCC|credit card|CL', product, ignore.case = T) ~ 'HDFC Retail', 
                            TRUE ~ as.character(lender)))


write.xlsx(Final_Collection, file = "C:\\R\\Collection\\Output\\Payment_file.xlsx")


axis_base<- Final_Collection %>% filter(lender %in% c('AXIS BANK'))
write.xlsx(axis_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_AXIS-", Sys.Date(), '.xlsx'))


# aye_base<- Final_Collection %>% filter(grepl('Aye',lender,ignore.case=TRUE))
# write.xlsx(aye_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_AYE-", Sys.Date(), '.xlsx'))

cap_base<- Final_Collection %>% filter(grepl('CAPITAL',lender,ignore.case=TRUE))
write.xlsx(cap_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_IDFC-", Sys.Date(), '.xlsx'))


CITI_CC_base<- Final_Collection %>% filter(grepl('CITI',lender,ignore.case=TRUE) & product %in% c('CC', 'CCC', 'CL'))
write.xlsx(CITI_CC_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_CITI_CC-", Sys.Date(), '.xlsx'))


FULL_base<- Final_Collection %>% filter(grepl('FULLERTON',lender,ignore.case=TRUE))
write.xlsx(FULL_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_FULLERTON-", Sys.Date(), '.xlsx'))


hdfc_CC_base<- Final_Collection %>% filter(hdfc_flag %in% c('HDFC CC') & product %in% c('CC', 'CCC', 'CL'))
write.xlsx(hdfc_CC_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_HDFC_CC-", Sys.Date(), '.xlsx'))

hdfc_Retail_base<- Final_Collection %>% filter(hdfc_flag %in% c('HDFC') & !product %in% c('CC', 'CCC', 'CL'))
write.xlsx(hdfc_Retail_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_HDFC_Retail-", Sys.Date(), '.xlsx'))

IDFC_base<- Final_Collection %>% filter(lender %in% c('CAPITAL FIRST'))
write.xlsx(IDFC_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_IDFC-", Sys.Date(), '.xlsx'))

ICICI_base<- Final_Collection %>% filter(lender %in% c('ICICI'))
write.xlsx(ICICI_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_ICICI-", Sys.Date(), '.xlsx'))

SBI_base<- Final_Collection %>% filter(lender %in% c('SBI Cards'))
write.xlsx(SBI_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_SBI-", Sys.Date(), '.xlsx'))

hero_base<- Final_Collection %>% filter(lender %in% c('Hero Fincorp Limited'))
write.xlsx(hero_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_Hero-", Sys.Date(), '.xlsx'))

abfl_base<- Final_Collection %>% filter(lender %in% c('Aditya Birla'))
write.xlsx(abfl_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_ABFL-", Sys.Date(), '.xlsx'))

scuf_base<- Final_Collection %>% filter(lender %in% c('Shriram City Union Finance Ltd'))
write.xlsx(scuf_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_Shriram-", Sys.Date(), '.xlsx'))

scb_base<- Final_Collection %>% filter(lender %in% c('SCB'))
write.xlsx(scb_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_SCB-", Sys.Date(), '.xlsx'))

kb_base<- Final_Collection %>% filter(lender %in% c('Krazybee Services Private Limited'))
write.xlsx(kb_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_KB-", Sys.Date(), '.xlsx'))


tvs_base<- Final_Collection %>% filter(lender %in% c('TVS'))
write.xlsx(tvs_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_TVS-", Sys.Date(), '.xlsx'))

CITI_base<- Final_Collection %>% filter(lender %in% c('CITI') & !product %in% c('CC', 'CCC', 'CL'))
write.xlsx(CITI_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_CITI-", Sys.Date(), '.xlsx'))

kp_base<- Final_Collection %>% filter(lender %in% c('KOTAK','Phoenix'))
write.xlsx(kp_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_Kotak&Phoenix-", Sys.Date(), '.xlsx'))

LNT_base<- Final_Collection %>% filter(lender %in% c('L & T Finance'))
write.xlsx(LNT_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_LNT-", Sys.Date(), '.xlsx'))


RBL_base<- Final_Collection %>% filter(lender %in% c('RBL'))
write.xlsx(RBL_base, file = paste0("C:\\R\\Collection\\Output\\Payment_file_RBL-", Sys.Date(), '.xlsx'))

write.xlsx(AAN_dump, file = paste0("C:\\R\\Collection\\Output\\AAN_dump-", Sys.Date(), '.xlsx'))
write.xlsx(RP_dump, file = paste0("C:\\R\\Collection\\Output\\RP_dump-", Sys.Date(), '.xlsx'))
