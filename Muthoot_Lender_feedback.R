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

Mdf1 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Muthoot.xlsx",sheet = 'AUG') %>% select(Mobile.Number,Location.feedback,Final.status)

# Mdf2 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Muthoot.xlsx",sheet = 'jun') %>% select(Mobile.Number,Location.feedback,Final.status)
# 
# Mdf3 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Muthoot.xlsx",sheet = 'july') %>% select(Mobile.Number,Location.feedback,Final.status)

Mdf2 <- read.xlsx("C:\\R\\Lender Feedback\\Input\\Muthoot.xlsx",sheet = 'sept')# %>% select(Mobile.Number,Feedback)


Mdf2$Mobile.Number <- as.numeric(Mdf2$Mobile.Number)

Mdf2$Feedback <- as.character(Mdf2$Feedback)


colnames(Mdf2)[which(names(Mdf2) == "Feedback")] <- "Final.status"

#names(Mdf2)

#df1 <- read.csv("data.csv", colClasses=c("numeric", "Date", "character"))
#df2 <- read.csv("data.csv", colClasses=c("numeric", "Date", "character"))

mudf1<-Mdf1 %>% full_join(Mdf2)



#mudf1<-Mdf3 %>% full_join(Mdf2)

#mudf2<-mudf1 %>% full_join(Mdf1)

#Muthoot_lender_MIS<-mudf2

Muthoot_lender_MIS<-mudf1

appos_map_Muthoot <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Muthootpl')
Muthoot_appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(grepl('Muthoot',name,ignore.case=TRUE)) %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################Muthoot LenderFeedBack

colnames(Muthoot_lender_MIS)[which(names(Muthoot_lender_MIS) == "Mobile.Number")] <- "mobile_number"

colnames(Muthoot_lender_MIS)[which(names(Muthoot_lender_MIS) == "Final.status")] <- "customer_status_name"

#Muthoot_lender_MIS$mobile_number <- as.numeric(Muthoot_lender_MIS$mobile_number)


Muthoot_lender_MIS$offer_application_number<-Muthoot_appos_dump$offer_application_number[match(Muthoot_lender_MIS$mobile_number, Muthoot_appos_dump$phone_home)]



Muthoot_lender_MIS$CURRENT_appops_status<-Muthoot_appos_dump$appops_status_code[match(Muthoot_lender_MIS$mobile_number, Muthoot_appos_dump$phone_home)]


Muthoot_lender_MIS$New_appops_status<-appos_map_Muthoot$New_Status[match(Muthoot_lender_MIS$customer_status_name, appos_map_Muthoot$CM_Status)]

Muthoot_lender_MIS$NEW_appops_description<-appos_map_Muthoot$Status_Description[match(Muthoot_lender_MIS$customer_status_name, appos_map_Muthoot$CM_Status)]


Muthoot_lender_MIS <-Muthoot_lender_MIS %>%
  mutate(Remark = case_when(
    #is.na(mobile_number) ~ 'Not in CRM',
    New_appops_status <= CURRENT_appops_status ~ 'Repeated_Feedback_cases',
    is.na(offer_application_number) ~ 'Not in CRM',
    (New_appops_status ==990) ~ '990',
    (New_appops_status ==380 & CURRENT_appops_status <380) ~ '380',
    (New_appops_status ==480 & CURRENT_appops_status <480) ~ '480',
    (New_appops_status ==580 & CURRENT_appops_status > 480) ~ '580'
    
  ))

Muthoot_lender_MIS$Remark <- ifelse(is.na(Muthoot_lender_MIS$Remark), Muthoot_lender_MIS$New_appops_status, Muthoot_lender_MIS$Remark)

Muthoot_lender_MIS$Upload_Remarks <- paste(Muthoot_lender_MIS$customer_status_name,'-',Muthoot_lender_MIS$NEW_appops_description)

colnames(Muthoot_lender_MIS)[which(names(Muthoot_lender_MIS) == "offer_application_number")] <- "Application_Number"

Muthoot_lender_MIS<-Muthoot_lender_MIS %>% mutate('Lender'="Muthoot")

write.xlsx(Muthoot_lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Revised_Muthoot_FB_File.xlsx")

Muthoot_upload<-Muthoot_lender_MIS %>% filter(!Remark %in% c('Repeated_Feedback_cases', 'Not_in_CRM'), !is.na(Application_Number))

Muthoot_upload<- Muthoot_upload %>% select(Application_Number) %>% 
  mutate(`App_Ops_Status`=Muthoot_upload$NEW_appops_description,`Bank_Feedback_Date`=thisdate,`Appointment_Date`="Nil",`Notes`=Muthoot_upload$Upload_Remarks,`Offer_Reference_Number`="",`Loan_Sanctioned_Disbursed_Amount`="",`Booking_Date`="",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")


write.xlsx(Muthoot_upload, file = "C:\\R\\Lender Feedback\\Output\\Muthoot_upload.xlsx")

#################################
