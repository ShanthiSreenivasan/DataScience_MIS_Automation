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
library(DBI)
library(RPostgreSQL)
library(RMySQL)
library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)
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

appos_dump <- read.csv("C:\\R\\Lender Feedback\\Input\\appopsdump.csv") %>% filter(name %in% c('CashE')) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

Lender_MIS<-read_xlsx("C:\\R\\Lender Feedback\\Input\\Lender_Feedback_MIS.xlsx", sheet='Sheet1') %>% filter(!Remarks %in% c('Repeated_Feedback_cases'))

names(appos_dump)

Lender_MIS$Current_APPOS<-appos_dump$appops_status_code[match(Lender_MIS$Application_Number, appos_dump$offer_application_number)]

Lender_MIS <-Lender_MIS %>%
  mutate(Status_flag = case_when(
    Remarks == Current_APPOS ~ 'Changed'
  ))

Lender_MIS$Status_flag[is.na(Lender_MIS$Status_flag)] <- 'Not changed'


# aggregate(cbind(count = Application_Number) ~ SKU, 
#           data = Lender_MIS, 
#           FUN = function(x){NROW(x)})
# n_row <- nrow(Lender_MIS %.% group_by(SKU))


#lender_maker <- Lender_MIS %>% group_by(SKU) %>% dplyr::summarise(`Total` = length(Application_Number)) %>% adorn_totals("row")

#checker_maker <- Lender_MIS %>% group_by(SKU) %>% dplyr::summarise(Total = length(Application_Number)) %>% adorn_totals("row")



write.csv(Lender_MIS, file = "C:\\R\\Lender Feedback\\Output\\Lender_Checker_MIS.csv")


# CashE_upload<- LK_lender_MIS_new_rej %>% filter(NEW_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,NEW_appops_description) %>% 
#   mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")
# 
# list_of_datasets <- list("FeedBack_update" = LK_lender_MIS_new_rej, "FeedBack_summary" = desc_LK_lender_MIS)
# write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\CashE_feedback.xlsx")
# 
# write.xlsx(CashE_upload, file = "C:\\R\\Lender Feedback\\CashE_upload.xlsx")


#####################################CHECKER MIS - mail sent


today=format(Sys.Date(), format="%d-%B-%Y")


filename<- "C:\\R\\Lender Feedback\\Output\\Lender_Checker_MIS.csv"


myMessage = paste0("Lender Feedback Uploaded Checker MIS- ",today,sep='')
sender <- "shanthi.s@creditmantri.com"

mail_LNT<-c("shanthi.s@creditmantri.com")
cclist<-c("shanthi.s@creditmantri.com")
# mail_LNT <- c("rakshith.thangaraj@creditmantri.com",
#                "sankarnarayanan@creditmantri.com","balamurugan.r@creditmantri.com",
#               "manikandan@creditmantri.com","pradeep.e@creditmantri.com","surya.s@creditmantri.com") 
#cclist <-c("Credit-Ops@creditmantri.com")

msg = '
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <style>
    table {font-family:  Verdana, Geneva, sans-serif; font-size: 11px;
    width = 100%; border: 1px solid black; border-collapse: collapse;
    text-align: center; padding: 5px;}
    th {height = 12px;background-color: #4CAF50;color: white;}
    td {background-color: #FFF;}
    </style>
  </head>
  <body>
    <h3> Uploaded Summary </h3>
    <p> ${print(xtable(lender_maker, digits = 0), type = \'html\')} </p><br>
    <h3> Current Status Summary </h3>
    <p> ${print(xtable(checker_maker, digits = 0), type = \'html\')} </p><br>
    
</body>
</html>'


if(file.exists(filename)){
  email <- send.mail(from = sender,
                     to = mail_LNT,
                     cc = c(cclist),
                     subject=myMessage,
                     html = TRUE,
                     inline = T,
                     body = str_interp(msg),
                     smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                                 user.name = "AKIA6IP74RHPWNJ7SJUL",
                                 passwd = "BPahpceHjt3FSbdHWybStClDcsZyYZYLP/JSKE50Tyd8" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}



