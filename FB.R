library(xtable)
library(purrr)
library(janitor)
library(scales)
library(tidyverse)
library(ggrepel)
library(lookup)
library(digest)
library(data.table)
#library(reshape2)
library(stringr)
library(httr)
require(stats)
library(mailR)


today=format(Sys.Date(), format="%d-%B-%Y")


##################################################
#View(CIS)
df2 <- read_xlsx("C:\\R\\Lender Feedback\\xls\\FB_lender_list.xlsx", sheet='MAIN')

df2<-df2 %>% select(2,3,4,5,6,8,9)

df3 <- read_xlsx("C:\\R\\Lender Feedback\\Output\\Final_FB_MIS.xlsx")


filename<-("C:\\R\\Lender Feedback\\Output\\Final_FB_MIS_dump.xlsx")

myMessage = paste0("Ops Referral Lender_Feedback_Update_MIS -",today,sep='')

sender <- "shanthi.s@creditmantri.com"

#mail_LNT<-c("shanthi.s@creditmantri.com")
#cclist<-c("shanthi.s@creditmantri.com")
mail_LNT <- c("hema@creditmantri.com",
              "Referral@creditmantri.com") 
cclist <-c("r.sudarshan@creditmantri.com","sankarnarayanan@creditmantri.com","rakshith.thangaraj@creditmantri.com","balamurugan.r@creditmantri.com","Ops-Referrals@creditmantri.com")

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
    <h3> Lender Feedback Summary </h3>
    <p> ${print(xtable(df2, digits = 0), type = \'html\')} </p><br>
    
    <h3> AOS Summary </h3>
    <p> ${print(xtable(df3, digits = 0), type = \'html\')} </p><br>
    
</body>
</html>'

#msg<-print(xtable(Overall_Panel,caption = "Find below the PA_Cost Panel-wise Summary"), type="html", caption.placement = "top")

if(file.exists(filename)){
  email <- send.mail(from = sender,
                     to = mail_LNT,
                     cc = c(cclist),
                     subject=myMessage,
                     html = TRUE,
                     inline = T,
                     body = str_interp(msg),
                     smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                                 user.name = "AKIA6IP74RHP36A2A7WY",
                                 passwd = "BMRvtdvA5TlHac3vFMtO3aTFAT8wXFVEod9ZkgoftvKk" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}



