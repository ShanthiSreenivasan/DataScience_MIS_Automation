rm(list=ls())
setwd('C:\\R\\S2L_Manual_Lenders\\Output')
today <- Sys.Date()

#load Library
library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
#library(DBI)
#library(RPostgreSQL)
#library(RMySQL)
#library(logging)
library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
#library(bit64)
#library(ggrepel)
#library(lookup)
#library(fs)
#library(htmlTable)
library(xtable)
library(yaml)

mail_list <- 'C:\\R\\S2L_Manual_Lenders\\Input\\MailMasterNew.xlsx'
mail_PNB <- read_excel(mail_list, sheet = "Sheet1")


mail <- list(mail_PNB)

#########################################

lapply(mail, function(x){
  
  x <- data.frame(x)
  lender <- x$File
  
  lapply(lender, function(y){
    
    row_set <- x[grepl(y, x$File),]
    sender <- 'referrals@creditmantri.com'
    
    recepients <- stringr::str_split(row_set$SendId[1], ",\\s*")[[1]]
    myMessage <- paste(row_set$SKU,row_set$Subject,today)
    msg <- paste('Hi Team,<br><br>Please find attached the list of new referrals for ',today,'.<br><br>Regards<br>CreditMantri Referral Team',sep='')
    
    cclist <- stringr::str_split(row_set$CC[1], ",\\s*")[[1]]
    
    filename <- paste('C:\\R\\S2L_Manual_Lenders\\Output\\',row_set$File,sep='')
    
    if(file.exists(filename)){ 
      email <- send.mail(from = sender,
                         to = recepients,
                         subject=myMessage,
                         cc=cclist,
                         html = TRUE,
                         inline = T,
                         body = str_interp(msg),
                         smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                                     user.name = "AKIAXJKXCWJGAXCHHIGY",
                                     passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
                         authenticate = TRUE,
                         attach.files = filename,
                send = TRUE)
    }
    
  })
  
  
  
})



      