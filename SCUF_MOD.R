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

setwd("C:\\R\\SCUF")


today_dat <- Sys.Date()


thisDate = Sys.Date()
if(!dir.exists(paste(".\\Input\\",thisDate,sep="")))
{
  dir.create(paste(".\\Input\\",thisDate,sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,sep=""))
} 



SCUF_dump <- read.csv("C:\\R\\SCUF\\Input\\Shriram Docs Upload MTD Cases.csv") %>% filter(applied_oic %in% c('System','andriodApp','PORTFOLIO') & current_oic %in% c('System','andriodApp','PORTFOLIO'))

SCUF_dump<-SCUF_dump %>% filter(document_status %in% c('completed'), current_aps %in% c(200), current_pos %in% c('PL-230','SBPL-230'), grepl('NA|New|WIP|CD|CL',current_crm, ignore.case = T) & !bs_status_shriram %in% c('Approved','Reject'))



live_df<-SCUF_dump %>% summarize(`Live`=n_distinct(phone_home, na.rm = TRUE))

assisted_dump <- read.xlsx("C:\\R\\SCUF\\Input\\Lender_sheet-INPUT.xlsx",sheet='Sheet1')

assisted_df = data.frame(assisted_dump) %>% summarise(Assisted=last(assisted_dump$Assisted))

scuf_df<-cbind(live_df,assisted_df) %>% mutate(`Total`=sum(live_df+assisted_df))

write.csv(scuf_df, file = paste0("C:\\R\\SCUF\\Output\\SCUF_MOD-", Sys.Date(), '.csv'))

today=format(Sys.Date(), format="%d-%B-%Y")


filename<-("C:\\R\\SCUF\\Input\\Shriram Docs Upload MTD Cases.csv")

myMessage = paste0("SCUF MOD MIS- ",today,sep='')
sender <- "Credit-Ops@creditmantri.com"

mail_LNT <- c("rakshith.thangaraj@creditmantri.com",
              "sankarnarayanan@creditmantri.com","balamurugan.r@creditmantri.com") 
cclist <-c("Credit-Ops@creditmantri.com")

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
    
    <h3> SCUF MOD Summary </h3>
    <p> ${print(xtable(scuf_df, digits = 0), type = \'html\')} </p><br>
    
    
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
                                 user.name = "AKIAXJKXCWJGAXCHHIGY",
                                 passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}




