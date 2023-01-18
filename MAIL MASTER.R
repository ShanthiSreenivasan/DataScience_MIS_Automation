#Remove existing list
rm(list=ls())

#Set working Directory
#setwd('Z:\\Amrish\\CIS')
Sys.setenv('JAVA_HOME'="C:\\Program Files\\Java\\jdk1.8.0_172/")
setwd(Sys.getenv('CIS NEW LENDER BATCH'))
library(dplyr)
library(readxl)
library(mailR)
library(stringi)
library(stringr)
library(XLConnect)


Send_Mail <- function(sender,recepients,myMessage,msg,cclist,filename= NULL){
  
send.mail(from = sender,
          to = c(recepients),
          subject = myMessage,
          html = TRUE,
          cc = c(cclist),
          inline = T,
          body = str_interp(msg),
          smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                      user.name = "AKIAXJKXCWJGAXCHHIGY",
                      passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
          authenticate = TRUE,
          attach.files =filename,
          send = TRUE)
}


#Shanker Mail
MailMaster <- read_excel('.\\Input\\MailMaster_Batch_New_zip.xlsx',sheet = 'Sankar')
#Sandra mails
MailMaster1 <- read_excel('.\\Input\\MailMaster_Batch_New_zip.xlsx',sheet = 'Sandra')
#Payment file
MailMaster2 <- read_excel('.\\Input\\MailMaster_Batch_New_zip.xlsx',sheet = 'Payments')
#new lender
MailMaster3 <- read_excel('.\\Input\\MailMaster_Batch_New_zip.xlsx',sheet = 'Others')

today <- format(Sys.Date(),"%Y-%m-%d")
Yesday <- format(Sys.Date()-1,"%Y-%m-%d")

#sankar
MailMaster_ZIP <- MailMaster %>% filter(grepl(".zip",`File Name`)) 
MailMaster_ZIP_1 <- MailMaster_ZIP %>% filter(!grepl("SBI Cards New cases",`File Name`))
MailMaster_ZIP_2 <- MailMaster_ZIP %>% filter(grepl("SBI Cards New cases",`File Name`))
MailMaster_XLSX <- MailMaster %>% filter(grepl(".xlsx",`File Name`))

#sandra & Others
MailMaster1_ZIP <- MailMaster1 %>% filter(grepl(".zip",`File Name`))
MailMaster1_XLSX <- MailMaster1 %>% filter(grepl(".xlsx",`File Name`))

ZIP_FILE <- rbind(MailMaster1_ZIP,MailMaster3)


#Sankar mails------------------------------------------------------------------------

#sankar
lapply(MailMaster_ZIP_1$`File Name`,function(x){
  
  #browser()
  path1 <-stringr::str_replace(paste('.\\Output\\',today,'\\Batch to lender New cases\\',x,sep=''),'.zip','.xlsx')
    if (file.exists(path1)) {
      myWorkbook <- XLConnect::loadWorkbook(path1)
    numberofsheets <- length(getSheets(myWorkbook))
    if(numberofsheets>1){
      
      data_file1 <- read_excel(path1,sheet = 1)
      data_file2 <- read_excel(path1,sheet = 2)
      count1<-nrow(data_file1)
      count2<-nrow(data_file2)
      
      if (count1>0) {
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_ZIP_1$To[MailMaster_ZIP_1$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_ZIP_1$Subject[MailMaster_ZIP_1$`File Name` == x]
        cclist = str_split(MailMaster_ZIP_1$cc[MailMaster_ZIP_1$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains rest of ",count2," LTD CM subscribed customer's & would require your assistance to provide us total amount dues towards complete bureau closure.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}
      }else{
      
      data_file1 <- read_excel(path1,sheet = 1)
      count1<-nrow(data_file1)
      
      if (count1>0) {
        
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_ZIP_1$To[MailMaster_ZIP_1$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_ZIP_1$Subject[MailMaster_ZIP_1$`File Name` == x]
        cclist = str_split(MailMaster_ZIP_1$cc[MailMaster_ZIP_1$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of ",count1,"recent subscribed customers who have an immediate intent to clear their dues. Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}}}
})

#SBI
lapply(MailMaster_ZIP_2$`File Name`,function(x){
  
  #browser()
  path1 <-stringr::str_replace(paste('.\\Output\\',today,'\\Batch to lender New cases\\',x,sep=''),'.zip','.xlsx')
  if (file.exists(path1)) {
    myWorkbook <- XLConnect::loadWorkbook(path1)
    numberofsheets <- length(getSheets(myWorkbook))
    if(numberofsheets>2){
      
      data_file1 <- read_excel(path1,sheet = 1)
      data_file2 <- read_excel(path1,sheet = 2)
      count1<-nrow(data_file1)
      count2<-nrow(data_file2)
      
      if (count1>0) {
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_ZIP_2$To[MailMaster_ZIP_2$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_ZIP_2$Subject[MailMaster_ZIP_2$`File Name` == x]
        cclist = str_split(MailMaster_ZIP_2$cc[MailMaster_ZIP_2$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains rest of ",count2," LTD CM subscribed customer's & would require your assistance to provide us total amount dues towards complete bureau closure.<br><br><b>Sheet 3:</b> Contains list of Payment files received till - ",Yesday,"<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}
    }else if (numberofsheets>1 & numberofsheets<3) {
      
      data_file1 <- read_excel(path1,sheet = 1)
      count1<-nrow(data_file1)
      
      if (count1>0) {
        
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_ZIP_2$To[MailMaster_ZIP_2$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_ZIP_2$Subject[MailMaster_ZIP_2$`File Name` == x]
        cclist = str_split(MailMaster_ZIP_2$cc[MailMaster_ZIP_2$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains list of Payment files received till - ",Yesday,"<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}}}
})

#EXCEL
lapply(MailMaster_XLSX$`File Name`,function(x){
  
  #browser()
  path1 <- paste('.\\Output\\',today,'\\Batch to lender New cases\\',x,sep='')
  if (file.exists(path1)) {
    myWorkbook <- XLConnect::loadWorkbook(path1)
    numberofsheets <- length(getSheets(myWorkbook))
    if(numberofsheets>1){
      
      data_file1 <- read_excel(path1,sheet = 1)
      data_file2 <- read_excel(path1,sheet = 2)
      count1<-nrow(data_file1)
      count2<-nrow(data_file2)
      
      if (count1>0) {
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_XLSX$To[MailMaster_XLSX$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_XLSX$Subject[MailMaster_XLSX$`File Name` == x]
        cclist = str_split(MailMaster_XLSX$cc[MailMaster_XLSX$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains rest of ",count2," LTD CM subscribed customer's & would require your assistance to provide us total amount dues towards complete bureau closure.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}
    }else{
      
      data_file1 <- read_excel(path1,sheet = 1)
      count1<-nrow(data_file1)
      
      if (count1>0) {
        
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster_XLSX$To[MailMaster_XLSX$`File Name` == x], ",\\s*")[[1]]
        myMessage <- MailMaster_XLSX$Subject[MailMaster_XLSX$`File Name` == x]
        cclist = str_split(MailMaster_XLSX$cc[MailMaster_XLSX$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of ",count1,"recent subscribed customers who have an immediate intent to clear their dues. Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}}}
})

#Payment files-----------------------------------------------------------------------

lapply(MailMaster2$`File Name`,function(x){
  
  
  #browser()
  # sender <- "referrals@creditmantri.com"
  sender <- 'Cislender@creditmantri.com'
  recepients <- str_split(MailMaster2$To[MailMaster2$`File Name` == x], ",\\s*")[[1]]
  myMessage <- paste(MailMaster2$Subject[MailMaster2$`File Name` == x]," ",Sys.Date() - 1,sep = '')
  cclist = str_split(MailMaster2$cc[MailMaster2$`File Name` == x], ",\\s*")[[1]]
  filename <- paste(".\\Output\\",today,"\\Payment file\\Zip_File\\",x,sep = "")
  msg <- paste("Hi All,<br><br>Please find the attached Payment file till - ",Sys.Date() - 1,".<br><br>Regards<br>Credit Mantri")
  
  if(file.exists(filename)){
    Mailing <- Send_Mail(sender,
                         recepients,
                         myMessage,
                         msg,
                         cclist,
                         filename)}
  
})

#Sandra & others---------------------------------------------------------------------

lapply(MailMaster1_XLSX$`File Name`,function(x){
  
  #browser()
  path1 <- paste('.\\Output\\',today,'\\Batch to lender New cases\\',x,sep='')
  if (file.exists(path1)) {
    myWorkbook <- XLConnect::loadWorkbook(path1)
    numberofsheets <- length(getSheets(myWorkbook))
    if(numberofsheets>1){
      
      data_file1 <- read_excel(path1,sheet = 1)
      data_file2 <- read_excel(path1,sheet = 2)
      count1<-nrow(data_file1)
      count2<-nrow(data_file2)
      
      if (count1>0) {
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster1_XLSX$To[MailMaster1_XLSX$`File Name` == x], ",\\s*")[[1]]
        myMessage <- paste(MailMaster1_XLSX$Subject[MailMaster1_XLSX$`File Name` == x]," ",Sys.Date(),sep = '')
        cclist = str_split(MailMaster1_XLSX$cc[MailMaster1_XLSX$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains rest of ",count2," LTD CM subscribed customer's & would require your assistance to provide us total amount dues towards complete bureau closure.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}
    }else{
      
      data_file1 <- read_excel(path1,sheet = 1)
      count1<-nrow(data_file1)
      
      if (count1>0) {
        
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(MailMaster1_XLSX$To[MailMaster1_XLSX$`File Name` == x], ",\\s*")[[1]]
        myMessage <- paste(MailMaster1_XLSX$Subject[MailMaster1_XLSX$`File Name` == x]," ",Sys.Date(),sep = '')
        cclist = str_split(MailMaster1_XLSX$cc[MailMaster1_XLSX$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of ",count1,"recent subscribed customers who have an immediate intent to clear their dues. Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}}}
})


lapply(ZIP_FILE$`File Name`,function(x){
  
  #browser()
  path1 <-stringr::str_replace(paste('.\\Output\\',today,'\\Batch to lender New cases\\',x,sep=''),'.zip','.xlsx')
  if (file.exists(path1)) {
    myWorkbook <- XLConnect::loadWorkbook(path1)
    numberofsheets <- length(getSheets(myWorkbook))
    if(numberofsheets>1){
      
      data_file1 <- read_excel(path1,sheet = 1)
      data_file2 <- read_excel(path1,sheet = 2)
      count1<-nrow(data_file1)
      count2<-nrow(data_file2)
      
      if (count1>0) {
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(ZIP_FILE$To[ZIP_FILE$`File Name` == x], ",\\s*")[[1]]
        myMessage <- paste(ZIP_FILE$Subject[ZIP_FILE$`File Name` == x]," ",Sys.Date(),sep = '')
        cclist = str_split(ZIP_FILE$cc[ZIP_FILE$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of cases awaiting total outstanding dues towards complete bureau closure. Basis priority excel is segregated in two sheets as highlighted below.<br><br><b>Sheet 1:</b> Contains list of ",count1," recent subscribed customers who have an immediate intent to clear their dues.  Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br><b>Sheet 2:</b> Contains rest of ",count2," LTD CM subscribed customer's & would require your assistance to provide us total amount dues towards complete bureau closure.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}
    }else{
      
      data_file1 <- read_excel(path1,sheet = 1)
      count1<-nrow(data_file1)
      
      if (count1>0) {
        
        sender <- 'Cislender@creditmantri.com'
        recepients <- str_split(ZIP_FILE$To[ZIP_FILE$`File Name` == x], ",\\s*")[[1]]
        myMessage <- paste(ZIP_FILE$Subject[ZIP_FILE$`File Name` == x]," ",Sys.Date(),sep = '')
        cclist = str_split(ZIP_FILE$cc[ZIP_FILE$`File Name` == x], ",\\s*")[[1]]
        filename <- paste(".\\Output\\",today,"\\Batch to lender New cases\\Zip_File\\",x,sep = "")
        msg <- paste("Hi All,<br><br>Attached excel contains list of ",count1,"recent subscribed customers who have an immediate intent to clear their dues. Kindly provide us the total amount dues towards complete bureau closure along with allocation on priority basis.<br><br>Regards<br>Credit Mantri")
        
        if(file.exists(filename)){
          Mailing <- Send_Mail(sender,
                               recepients,
                               myMessage,
                               msg,
                               cclist,
                               filename)}}}}
})





