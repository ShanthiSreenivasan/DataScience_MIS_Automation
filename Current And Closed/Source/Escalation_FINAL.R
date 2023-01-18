rm(list = ls())
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_172')
library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(stringr)
library(mailR)


setwd("C:\\R\\Current And Closed")
getwd() 
source('.\\Functionfile.R')

today_dat <- Sys.Date()

date1 <-format(Sys.Date()-1,"%d%m")

thisDate = Sys.Date()

if(!dir.exists(paste(".\\Output\\",thisDate,sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,"\\Escalation",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\Escalation",sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,"\\o2CEscalation",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\o2CEscalation",sep=""))
} 

#Read New cases dump --------------------------------


dump_path <- paste(".\\Input\\cis_dump_new_cases.csv",sep='')

dump <- fread(dump_path,na.strings = c('',NA))

dump <- dump[!is.na(dump$email_id),]
dump$CISserialno <- as.numeric(dump$CISserialno)

Esc1_data <-paste("./Input/ESCALATION_EMAIL_1_",date1,".CSV",sep = '')

Esc1 <-fread(Esc1_data) %>% select('lead_id') 
Esc1$lead_id <-as.numeric(Esc1$lead_id)

Esc2_data <-paste("./Input/ESCALATION_EMAIL_2_",date1,".CSV",sep = '')

Esc2 <-fread(Esc2_data) %>% select('lead_id')
Esc2$lead_id <-as.numeric(Esc2$lead_id)

Esc3_data <-paste("./Input/ESCALATION_EMAIL_3_",date1,".CSV",sep = '')

Esc3 <-fread(Esc3_data) %>% select('lead_id') 
Esc3$lead_id <-as.numeric(Esc3$lead_id)

fEsc_data <-paste("./Input/FINAL_ESCALATION_",date1,".CSV",sep = '')

fEsc <-fread(fEsc_data) %>% select('lead_id') 
fEsc$lead_id <-as.numeric(fEsc$lead_id)




o2CEsc1_data <-paste("./Input/02C_ESCALATION_EMAIL_1_",date1,".CSV",sep = '')

o2CEsc1 <-fread(o2CEsc1_data) %>% select('lead_id') 
o2CEsc1$lead_id <-as.numeric(o2CEsc1$lead_id)

o2CEsc2_data <-paste("./Input/02C_ESCALATION_EMAIL_2_",date1,".CSV",sep = '')

o2CEsc2 <-fread(o2CEsc2_data) %>% select('lead_id')
o2CEsc2$lead_id <-as.numeric(o2CEsc2$lead_id)

o2CEsc3_data <-paste("./Input/02C_ESCALATION_EMAIL_2_",date1,".CSV",sep = '')

o2CEsc3 <-fread(o2CEsc3_data) %>% select('lead_id') 
o2CEsc3$lead_id <-as.numeric(o2CEsc3$lead_id)


fo2CEsc_data <-paste("./Input/02C_FINAL_ESCALATION_",date1,".CSV",sep = '')

fo2CEsc <-fread(fo2CEsc_data) %>% select('lead_id') 
fo2CEsc$lead_id <-as.numeric(fo2CEsc$lead_id)

#Add Row -------------------------------------------------------------------------------

dump$Customer_Name=paste(dump$first_name,dump$last_name)
dump$Date=dump$subscription_date
dump$Phone_Number=dump$phone_home
dump$Address=dump$city
dump$Product_Status=dump$prodstatus
dump$Lender_Name=dump$lender_name
dump$Account_type=dump$account_type
dump$Email_Id=dump$email_id
dump$Account_status=dump$account_status

dump <- dump %>% mutate(Flag="")


#Escalation_1 ----------------------------------------------------------------------



Escalation_1 <- dump %>% filter(CISserialno %in% Esc1$lead_id)


Escalation_1<-Escalation_1%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_esc1 <-Esc1 %>% filter(!lead_id  %in% Escalation_1$CISserialno)

if (nrow(Escalation_1)>0) {
  
  
  
  Escalation_1 <- Escalation_1 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                          'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-Escalation-1.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Esc_dump1 <- left_join(Escalation_1, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Escalation_no_mail_id_1 <- Esc_dump1[is.na(Esc_dump1$Email),]
  Escalation_dump1 <- Esc_dump1[!is.na(Esc_dump1$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- Escalation_dump1
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(Escalation_dump1$CISserialno, function(lead){
    
    #browser()
    accounts <- Escalation_dump1[Escalation_dump1$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg:Escalation Email 1 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    # email_body <- str_interp(' Dear ${accounts$Customer_Name}, <br/> <br/>Ref: Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/> We had sent you four emails a few days ago regarding your subscription to the Credit Improvement Service with CreditMantri and requested you to forward the same to your lender. Since we have not heard from the lender yet, we are following-up to remind you. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ${accounts$Email} from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com<br/><br/>Please forward us any responses that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/><br/>---------------------------------------Cut Here---------------------------------------
    #                          <br/><br/>Dear Sir / Madam, <br/> <br/>Reg:Escalation Email 1 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No} <br/><br/> This is with regards to above mentioned account that has been reported with a ${accounts$Account_status} status in the repayment history with the Equifax Credit Bureau. <br/><br/>I have tried to contact your service desk via email to understand what needs to be done in order to rectify this negative reporting, but have not received any response from them despite my numerous reminders. I am therefore bringing this to your attention anticipating a speedy revert.<br/><br/>I request you to let me know the status of the account and the outstanding amount, if any, to be paid by me OR let me know the best way to change the status of the account to CLOSED and STANDARD in my credit report.<br/><br/>If I do not receive any response from you I will be forced to report the issue to the Banking Ombudsmen in order to seek resolution.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name} 
    #                          <br/> <br/>---------------------------------------Cut Here---------------------------------------
    #                          <br/> <br/>Kindly forward any response/feedback that you may receive from them. <br/> <br/>Thanks and Regards<br/>${accounts$Customer_Name}')
    # 
    #New content -----
    email_body <- str_interp('Dear ${accounts$Customer_Name},<br/> <br/>
                           
                           Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>
                           
                           We had sent you four emails a few days ago regarding your subscription to the Credit Improvement Service with CreditMantri and requested you to forward the same to your lender. Since we have not heard from the lender yet, we are following-up to remind you. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender (${accounts$Email}) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com<br/><br/>
                           
                           Please forward us any responses that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                             
                           Dear Sir / Madam:,
                           
                           <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}
                           
                           <br/> <br/>This is with regards to the above-mentioned account that has been reported with a 60-89 days past due status in the repayment history with the Equifax Credit Bureau.
                           
                           I have tried to contact your service desk via email to understand what needs to be done in order to rectify this negative reporting, but have not received any response from them despite my numerous reminders. I am therefore bringing this to your attention anticipating a speedy revert.
                           
                           I request you to let me know the status of the account and the outstanding amount, if any, to be paid by me OR let me know the best way to change the status of the account to CLOSED and STANDARD in my credit report.
                           
                           If I do not receive any response from you I will be forced to report the issue to the Banking Ombudsman in order to seek resolution.<br/><br/>
                           
                           Thanks and Regards<br/>
                           ${accounts$Customer_Name}<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                           
                           Kindly forward any response/feedback that you may receive from them.<br/><br/>
                           
                           Thanks and Regards<br/>
                          Team CreditMantri')
    
    
    #recipients <- accounts$Email
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("Escalation one -",accounts$CISserialno))
    
  })
  
}else{
  
  
  Escalation_1 <- Escalation_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                         'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                         'LDB','Flag')
  
  Escalation_no_mail_id_1 <-Escalation_1
  Escalation_dump1 <-Escalation_1
  
  NewLeads <- Escalation_1
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


#Escalation_2 ----------------------------------------------------------------------



Escalation_2 <- dump %>% filter(CISserialno %in% Esc2$lead_id)

Escalation_2<-Escalation_2%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))


LeadsNotFound_esc2 <-Esc2 %>% filter(!lead_id  %in% Escalation_2$CISserialno)


if (nrow(Escalation_2)>0) {
  
  
  Escalation_2 <- Escalation_2 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                          'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-Escalation-2.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Esc_dump2 <- left_join(Escalation_2, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Escalation_no_mail_id_2 <- Esc_dump2[is.na(Esc_dump2$Email),]
  Escalation_dump2 <- Esc_dump2[!is.na(Esc_dump2$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- Escalation_dump2
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
  lapply(Escalation_dump2$CISserialno, function(lead){
    
    #browser()
    accounts <- Escalation_dump2[Escalation_dump2$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : Escalation Email 2 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    # email_body <- str_interp(' Dear ${accounts$Customer_Name} <br/> <br/>Ref: Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name} 
    #                          <br/> <br/>This is our 2nd reminder regarding the Credit Improvement Service that you had subscribed to with CreditMantri and we have sent to the lender. We have not yet received confirmation from the lender regarding sending out the email to your lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ${accounts$Email}from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com
    #                          <br/> <br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
    #                          <br/> <br/><br/>---------------------------------------Cut Here--------------------------------------
    #                          <br/> <br/>Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>I had earlier sent four emails requesting for an update however have still not received it.
    #                          <br/> <br/>I am herewith following up to remind you that on scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD.
    #                          <br/> <br/>I have tried to contact your office about the status of this account but have not been able to get any response. Hence, I am bringing this to your attention and expecting a speedy response.
    #                          <br/> <br/>If I do not receive any response from you, I will be forced to take the issue to the Banking Ombudsman for resolution.
    #                          <br/> <br/>Thanks and Regards<br/>${accounts$Customer_Name}
    #                          <br/> <br/><br/>---------------------------------------Cut Here--------------------------------------
    #                          <br/> <br/>Kindly forward any response/feedback that you may receive from them.<br/> <br/>Thanks and Regards<br/>CreditMantri Team')
    # 
    # #recipients <- accounts$Email
    
    email_body <- str_interp('Dear ${accounts$Customer_Name}, <br/> <br/>

                           Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}
                           
                           <br/> <br/>This is our 2nd reminder regarding the Credit Improvement Service that you had subscribed to with CreditMantri and we have sent to the lender. We have not yet received confirmation from the lender regarding sending out the email to your lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender (${accounts$Email}) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com
                           
                           <br/> <br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
                           
                           <br/> <br/>---------------------------------------Cut Here---------------------------------------
                           
                           <br/> <br/>Dear Sir/Madam,
                           
                           <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}
                           
                           <br/> <br/>I had earlier sent four emails requesting for an update however have still not received it.
                           
                           I am herewith following up to remind you that on scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD.
                           
                           I have tried to contact your office about the status of this account but have not been able to get any response. Hence, I am bringing this to your attention and expecting a speedy response.
                           
                           If I do not receive any response from you, I will be forced to take the issue to the Banking Ombudsman for resolution.
                           
                           Thanks and Regards<br/>
                           ${accounts$Customer_Name}<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                           
                           Kindly forward any response/feedback that you may receive from them.<br/><br/>
                           
                           Thanks and Regards<br/>
                           Team CreditMantri')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("Escalation two -",accounts$CISserialno))
    
  })
  
  
  
}else{
  
  
  
  Escalation_2 <- Escalation_2 %>% mutate(Flag='')%>%  
    select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
           'LDB','Flag')
  Escalation_no_mail_id_2 <-Escalation_2
  Escalation_dump2 <-Escalation_2
  
  NewLeads <- Escalation_2
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}


#Escalation_3 ----------------------------------------------------------------------



Escalation_3 <- dump %>% filter(CISserialno %in% Esc3$lead_id)


Escalation_3<-Escalation_3%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_esc3 <-Esc3 %>% filter(!lead_id  %in% Escalation_3$CISserialno)

if (nrow(Escalation_3)>0) {
  
  
  
  Escalation_3 <- Escalation_3 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                          'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-Escalation-3.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Esc_dump3 <- left_join(Escalation_3, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Escalation_no_mail_id_3 <- Esc_dump3[is.na(Esc_dump3$Email),]
  Escalation_dump3 <- Esc_dump3[!is.na(Esc_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- Escalation_dump3
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(Escalation_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- Escalation_dump3[Escalation_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : Escalation Email 3 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    # email_body <- str_interp('Dear ${accounts$Customer_Name}<br/> <br/>Ref:Subscription to our Credit Improvement Service for your account with${accounts$Lender_Name}<br/><br/>This is our 3rd reminder to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from the lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.<br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.<br/><br/><br/>---------------------------------------Cut Here--------------------------------------<br/> <br/>
    #                          Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>This is my sixth reminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 60-89 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.
    #                          <br/><br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure else I will be forced to approach the RBI and Banking ombudsman seeking intervention.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name}
    #                          <br/><br/>---------------------------------------Cut Here--------------------------------------<br/><br/>Kindly forward any response/feedback that you may receive from them.<br/><br/>Thanks and Regards<br/>CreditMantri Team')
    # 
    #recipients <- accounts$Email
    
    email_body <- str_interp('Dear ${accounts$Customer_Name},

                           <br/> <br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}
                           
                           <br/> <br/>This is our 3rd reminder to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from the lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.
                           
                           <br/> <br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
                           
                           <br/> <br/>---------------------------------------Cut Here---------------------------------------
                           
                           <br/> <br/>Dear Sir/Madam,
                           
                           <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}
                           
                           <br/> <br/>This is my sixth reminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 60-89 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.
                           
                           <br/> <br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure else I will be forced to approach the RBI and Banking ombudsman seeking intervention.
                           
                           Thanks and Regards<br/>
                           ${accounts$Customer_Name}<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                           
                           Kindly forward any response/feedback that you may receive from them.<br/><br/>
                           
                           Thanks and Regards<br/>
                           Team CreditMantri')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("Escalation three -",accounts$CISserialno))
    
  })
  
  
}else{
  
  
  
  Escalation_3 <- Escalation_3 %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  Escalation_no_mail_id_3 <-Escalation_3
  Escalation_dump3 <-Escalation_3
  
  NewLeads <- Escalation_3
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}




#FINAL Escaltion------------------------------------------------------------------------------------------

FEscalation <- dump %>% filter(CISserialno %in% fEsc$lead_id)

FEscalation<-FEscalation%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_fesc <-fEsc %>% filter(!lead_id  %in% FEscalation$CISserialno)

if (nrow(FEscalation)>0) {
  
  
  
  FEscalation <- FEscalation %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                        'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                        'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-Closure.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  fEsc_dump3 <- left_join(FEscalation, product_family,
                          by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  FEscalation_no_mail_id_3 <- fEsc_dump3[is.na(fEsc_dump3$Email),]
  FEscalation_dump3 <- fEsc_dump3[!is.na(fEsc_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- FEscalation_dump3
  FilePath <- paste('./Output/',thisDate,"\\Escalation/FINAL Escalation Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(FEscalation_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- FEscalation_dump3[FEscalation_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : Final Escalation Email - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    #   email_body <- str_interp('Dear ${accounts$Customer_Name}<br/> <br/>Ref:Subscription to our Credit Improvement Service for your account with${accounts$Lender_Name}<br/><br/>This is our 3rd reminder to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from the lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.<br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.<br/><br/><br/>---------------------------------------Cut Here--------------------------------------<br/> <br/>
    #                            Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>This is my sixth reminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 60-89 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.
    #                            <br/><br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure else I will be forced to approach the RBI and Banking ombudsman seeking intervention.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name}
    #                            <br/><br/>---------------------------------------Cut Here--------------------------------------<br/><br/>Kindly forward any response/feedback that you may receive from them.<br/><br/>Thanks and Regards<br/>CreditMantri Team')
    # 
    # #recipients <- accounts$Email
    
    email_body <- str_interp('Dear ${accounts$Customer_Name},

                             <br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}
                             
                             <br/><br/>We have sent many reminders to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.
                             
                             <br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
                             
                             <br/><br/>---------------------------------------Cut Here---------------------------------------
                             
                             <br/><br/>Dear Sir/Madam,
                             
                             <br/><br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}
                             
                             <br/><br/>This is my seventh reminder about the email that I had sent earlier too and since I have not heard from you, I have not received any response to my previous mails to you. Request you to please help with urgent closure.  I am bringing this to your attention and also expect a reply within next one week failing which I will be left with no choice than to escalate further to ombudsman and related authorities.
                             
                             Thanks and Regards<br/>
                           ${accounts$Customer_Name}<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                           
                           Kindly forward any response/feedback that you may receive from them.<br/><br/>
                           
                           Thanks and Regards<br/>
                          Team CreditMantri')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("Escalation three -",accounts$CISserialno))
    
  })
  
  
}else{
  
  
  
  FEscalation <- FEscalation %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                         'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                         'LDB','Flag')
  
  FEscalation_no_mail_id_3 <-FEscalation
  FEscalation_dump3 <-FEscalation
  
  NewLeads <- FEscalation
  FilePath <- paste('./Output/',thisDate,"\\Escalation/FINAL Escalation Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


#o2CEscalation_1 ----------------------------------------------------------------------



o2CEscalation_1 <- dump %>% filter(CISserialno %in% o2CEsc1$lead_id)

o2CEscalation_1<-o2CEscalation_1%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_o2CEsc1 <-o2CEsc1 %>% filter(!lead_id  %in% o2CEscalation_1$CISserialno)

if (nrow(o2CEscalation_1)>0) {
  
  
  
  o2CEscalation_1 <- o2CEscalation_1 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2CEscalation-1.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CEsc_dump1 <- left_join(o2CEscalation_1, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CEscalation_no_mail_id_1 <- o2CEsc_dump1[is.na(o2CEsc_dump1$Email),]
  o2CEscalation_dump1 <- o2CEsc_dump1[!is.na(o2CEsc_dump1$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- o2CEscalation_dump1
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(o2CEscalation_dump1$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CEscalation_dump1[o2CEscalation_dump1$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg:02CEscalation Email 1 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    email_body <- str_interp(' Dear ${accounts$Customer_Name}, <br/> <br/>Ref: Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/> We had sent you four emails a few days ago regarding your subscription to the Credit Improvement Service with CreditMantri and requested you to forward the same to your lender. Since we have not heard from the lender yet, we are following-up to remind you. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ${accounts$Email} from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com <br/><br/>Please forward us any responses that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/><br/>---------------------------------------Cut Here---------------------------------------
                           <br/><br/>Dear Sir / Madam, <br/> <br/>Reg:o2CEscalation Email 1 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No} <br/><br/> This is with regards to above mentioned account that has been reported with a ${accounts$Account_status} status in the repayment history with the Equifax Credit Bureau. <br/><br/>I have tried to contact your service desk via email to understand what needs to be done in order to rectify this negative reporting, but have not received any response from them despite my numerous reminders. I am therefore bringing this to your attention anticipating a speedy revert.<br/><br/>I request you to let me know the status of the account and the outstanding amount, if any, to be paid by me OR let me know the best way to change the status of the account to CLOSED and STANDARD in my credit report.<br/><br/>If I do not receive any response from you I will be forced to report the issue to the Banking Ombudsmen in order to seek resolution.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name} 
                           <br/> <br/>---------------------------------------Cut Here---------------------------------------
                           <br/> <br/>Kindly forward any response/feedback that you may receive from them. <br/> <br/>Thanks and Regards<br/>Team CreditMantri')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("o2CEscalation one -",accounts$CISserialno))
    
  })
  
}else{
  
  
  o2CEscalation_1 <- o2CEscalation_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                               'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                               'LDB','Flag')
  
  o2CEscalation_no_mail_id_1 <-o2CEscalation_1
  o2CEscalation_dump1 <-o2CEscalation_1
  
  NewLeads <- o2CEscalation_1
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


#o2CEscalation_2 ----------------------------------------------------------------------



o2CEscalation_2 <- dump %>% filter(CISserialno %in% o2CEsc2$lead_id)

o2CEscalation_2<-o2CEscalation_2%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_o2CEsc2 <-o2CEsc2 %>% filter(!lead_id  %in% o2CEscalation_2$CISserialno)


if (nrow(o2CEscalation_2)>0) {
  
  
  o2CEscalation_2 <- o2CEscalation_2 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2CEscalation-2.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CEsc_dump2 <- left_join(o2CEscalation_2, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CEscalation_no_mail_id_2 <- o2CEsc_dump2[is.na(o2CEsc_dump2$Email),]
  o2CEscalation_dump2 <- o2CEsc_dump2[!is.na(o2CEsc_dump2$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- o2CEscalation_dump2
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
  lapply(o2CEscalation_dump2$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CEscalation_dump2[o2CEscalation_dump2$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : 02CEscalation Email 2 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    email_body <- str_interp(' Dear ${accounts$Customer_Name} <br/> <br/>Ref: Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name} 
                           <br/> <br/>This is our 2nd reminder regarding the Credit Improvement Service that you had subscribed to with CreditMantri and we have sent to the lender. We have not yet received confirmation from the lender regarding sending out the email to your lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender (${accounts$Email}) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com
                           <br/> <br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
                           <br/> <br/><br/>---------------------------------------Cut Here--------------------------------------
                           <br/> <br/>Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>I had earlier sent four emails requesting for an update however have still not received it.
                           <br/> <br/>I am herewith following up to remind you that on scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD.
                           <br/> <br/>I have tried to contact your office about the status of this account but have not been able to get any response. Hence, I am bringing this to your attention and expecting a speedy response.
                           <br/> <br/>If I do not receive any response from you, I will be forced to take the issue to the Banking Ombudsman for resolution.
                           <br/> <br/>Thanks and Regards<br/>${accounts$Customer_Name}
                           <br/> <br/><br/>---------------------------------------Cut Here--------------------------------------
                           <br/> <br/>Kindly forward any response/feedback that you may receive from them.<br/> <br/>Thanks and Regards<br/>CreditMantri Team')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("o2CEscalation two -",accounts$CISserialno))
    
  })
  
  
  
}else{
  
  
  
  o2CEscalation_2 <- o2CEscalation_2 %>% mutate(Flag='')%>%  
    select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
           'LDB','Flag')
  o2CEscalation_no_mail_id_2 <-o2CEscalation_2
  o2CEscalation_dump2 <-o2CEscalation_2
  
  NewLeads <- o2CEscalation_2
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}


#o2CEscalation_3 ----------------------------------------------------------------------



o2CEscalation_3 <- dump %>% filter(CISserialno %in% o2CEsc3$lead_id)

o2CEscalation_3<-o2CEscalation_3%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_o2CEsc3 <-o2CEsc3 %>% filter(!lead_id  %in% o2CEscalation_3$CISserialno)

if (nrow(o2CEscalation_3)>0) {
  
  
  
  o2CEscalation_3 <- o2CEscalation_3 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2CEscalation-3.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CEsc_dump3 <- left_join(o2CEscalation_3, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CEscalation_no_mail_id_3 <- o2CEsc_dump3[is.na(o2CEsc_dump3$Email),]
  o2CEscalation_dump3 <- o2CEsc_dump3[!is.na(o2CEsc_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- o2CEscalation_dump3
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(o2CEscalation_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CEscalation_dump3[o2CEscalation_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : 02CEscalation Email 3 - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    email_body <- str_interp('Dear ${accounts$Customer_Name}<br/> <br/>Ref:Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>This is our 3rd reminder to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from the lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.<br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.<br/><br/><br/>---------------------------------------Cut Here--------------------------------------<br/> <br/>
                           Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>This is my sixth reminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 60-89 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.
                           <br/><br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure else I will be forced to approach the RBI and Banking ombudsman seeking intervention.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name}
                           <br/><br/>---------------------------------------Cut Here--------------------------------------<br/><br/>Kindly forward any response/feedback that you may receive from them.<br/><br/>Thanks and Regards<br/>CreditMantri Team')
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("o2CEscalation three -",accounts$CISserialno))
    
  })
  
  
}else{
  
  
  
  o2CEscalation_3 <- o2CEscalation_3 %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                 'LDB','Flag')
  
  o2CEscalation_no_mail_id_3 <-o2CEscalation_3
  o2CEscalation_dump3 <-o2CEscalation_3
  
  NewLeads <- o2CEscalation_3
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}



#FINAL o2CEscaltion------------------------------------------------------------------------------------------

Fo2CEscalation <- dump %>% filter(CISserialno %in% fo2CEsc$lead_id)

Fo2CEscalation<-Fo2CEscalation%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_fo2CEsc <-fo2CEsc %>% filter(!lead_id  %in% Fo2CEscalation$CISserialno)

if (nrow(Fo2CEscalation)>0) {
  
  
  
  Fo2CEscalation <- Fo2CEscalation %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                              'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                              'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2CClosure.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  fo2CEsc_dump3 <- left_join(Fo2CEscalation, product_family,
                             by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Fo2CEscalation_no_mail_id_3 <- fo2CEsc_dump3[is.na(fo2CEsc_dump3$Email),]
  Fo2CEscalation_dump3 <- fo2CEsc_dump3[!is.na(fo2CEsc_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  NewLeads <- Fo2CEscalation_dump3
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/FINAL o2CEscalation Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  lapply(Fo2CEscalation_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- Fo2CEscalation_dump3[Fo2CEscalation_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    cc_recipients <- c('ciscare@creditmantri.com',accounts$Email_Id)
    
    #bcc_Recipinents <- 'Ciscare@creditmantri.com'
    
    message  <- str_interp('Reg : Final 02CEscalation Email - ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    #   email_body <- str_interp('Dear ${accounts$Customer_Name}<br/> <br/>Ref:Subscription to our Credit Improvement Service for your account with${accounts$Lender_Name}<br/><br/>This is our 3rd reminder to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from the lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purposes, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.<br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.<br/><br/><br/>---------------------------------------Cut Here--------------------------------------<br/> <br/>
    #                            Sir / Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>This is my sixth reminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 60-89 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.
    #                            <br/><br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure else I will be forced to approach the RBI and Banking ombudsman seeking intervention.<br/><br/>Thanks and Regards<br/>${accounts$Customer_Name}
    #                            <br/><br/>---------------------------------------Cut Here--------------------------------------<br/><br/>Kindly forward any response/feedback that you may receive from them.<br/><br/>Thanks and Regards<br/>CreditMantri Team')
    # 
    # #recipients <- accounts$Email
    
    email_body <- str_interp('Dear ${accounts$Customer_Name},

                             <br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}
                             
                             <br/><br/>We have sent many reminders to the lender regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.
                             
                             <br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.
                             
                             <br/><br/>---------------------------------------Cut Here---------------------------------------
                             
                             <br/><br/>Dear Sir/Madam,
                             
                             <br/><br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}
                             
                             <br/><br/>This is my seventh reminder about the email that I had sent earlier too and since I have not heard from you, I have not received any response to my previous mails to you. Request you to please help with urgent closure.  I am bringing this to your attention and also expect a reply within next one week failing which I will be left with no choice than to o2CEscalate further to ombudsman and related authorities.
                             
                             Thanks and Regards<br/>
                           ${accounts$Customer_Name}<br/><br/>
                           
                           ---------------------------------------Cut Here---------------------------------------<br/><br/>
                           
                           Kindly forward any response/feedback that you may receive from them.<br/><br/>
                           
                           Thanks and Regards<br/>
                          Team CreditMantri')
    
    
    mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body,filename= NULL)
    
    print(paste("o2CEscalation three -",accounts$CISserialno))
    
  })
  
  
}else{
  
  
  
  Fo2CEscalation <- Fo2CEscalation %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                               'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                               'LDB','Flag')
  
  Fo2CEscalation_no_mail_id_3 <-Fo2CEscalation
  Fo2CEscalation_dump3 <-Fo2CEscalation
  
  NewLeads <- Fo2CEscalation
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/FINAL o2CEscalation Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}




#unsent Email ----------------------------------------------------------------------------------------

wb <- createWorkbook()
addWorksheet(wb,"Escalation-1")
addWorksheet(wb,"Escalation-2")
addWorksheet(wb,"Escalation-3")
addWorksheet(wb,"FINAL Escalation")

hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"Escalation-1",cols = 1:ncol(Escalation_no_mail_id_1),widths = 15)
writeData(wb,"Escalation-1",Escalation_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-2",cols = 1:ncol(Escalation_no_mail_id_2),widths = 15)
writeData(wb,"Escalation-2",Escalation_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-3",cols = 1:ncol(Escalation_no_mail_id_3),widths = 15)
writeData(wb,"Escalation-3",Escalation_no_mail_id_3,borders = "all",headerStyle = hs1)


setColWidths(wb,"FINAL Escalation",cols = 1:ncol(FEscalation_no_mail_id_3),widths = 15)
writeData(wb,"FINAL Escalation",FEscalation_no_mail_id_3,borders = "all",headerStyle = hs1)

path <- paste('./Output/',thisDate,"\\Escalation/Unsent Emails on - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)

##################o2C

wb <- createWorkbook()

addWorksheet(wb,"o2CEscalation-1")
addWorksheet(wb,"o2CEscalation-2")
addWorksheet(wb,"o2CEscalation-3")
addWorksheet(wb,"FINAL o2CEscalation")

hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"o2CEscalation-1",cols = 1:ncol(o2CEscalation_no_mail_id_1),widths = 15)
writeData(wb,"o2CEscalation-1",o2CEscalation_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-2",cols = 1:ncol(o2CEscalation_no_mail_id_2),widths = 15)
writeData(wb,"o2CEscalation-2",o2CEscalation_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-3",cols = 1:ncol(o2CEscalation_no_mail_id_3),widths = 15)
writeData(wb,"o2CEscalation-3",o2CEscalation_no_mail_id_3,borders = "all",headerStyle = hs1)

setColWidths(wb,"FINAL o2CEscalation",cols = 1:ncol(Fo2CEscalation_no_mail_id_3),widths = 15)
writeData(wb,"FINAL o2CEscalation",Fo2CEscalation_no_mail_id_3,borders = "all",headerStyle = hs1)


o2CEsc_path2 <- paste('./Output/',thisDate,"\\o2CEscalation/Unsent Emails on - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,o2CEsc_path2,overwrite = T)




#No leads found in Crm---------------------------------------------------------------


wb <- createWorkbook()

addWorksheet(wb,"Escalation-1")
addWorksheet(wb,"Escalation-2")
addWorksheet(wb,"Escalation-3")
addWorksheet(wb,"FINAL Escalation")

hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"Escalation-1",cols = 1:ncol(LeadsNotFound_esc1),widths = 15)
writeData(wb,"Escalation-1",LeadsNotFound_esc1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-2",cols = 1:ncol(LeadsNotFound_esc2),widths = 15)
writeData(wb,"Escalation-2",LeadsNotFound_esc2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-3",cols = 1:ncol(LeadsNotFound_esc3),widths = 15)
writeData(wb,"Escalation-3",LeadsNotFound_esc3,borders = "all",headerStyle = hs1)

setColWidths(wb,"FINAL Escalation",cols = 1:ncol(LeadsNotFound_fesc),widths = 15)
writeData(wb,"FINAL Escalation",LeadsNotFound_fesc,borders = "all",headerStyle = hs1)

path <- paste('./Output/',thisDate,"\\Escalation/No Leads found in CRM - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)

#######################o2c

wb <- createWorkbook()

addWorksheet(wb,"o2CEscalation-1")
addWorksheet(wb,"o2CEscalation-2")
addWorksheet(wb,"o2CEscalation-3")

hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"o2CEscalation-1",cols = 1:ncol(LeadsNotFound_o2CEsc1),widths = 15)
writeData(wb,"o2CEscalation-1",LeadsNotFound_o2CEsc1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-2",cols = 1:ncol(LeadsNotFound_o2CEsc2),widths = 15)
writeData(wb,"o2CEscalation-2",LeadsNotFound_o2CEsc2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-3",cols = 1:ncol(LeadsNotFound_o2CEsc3),widths = 15)
writeData(wb,"o2CEscalation-3",LeadsNotFound_o2CEsc3,borders = "all",headerStyle = hs1)

setColWidths(wb,"FINAL Escalation",cols = 1:ncol(LeadsNotFound_fo2CEsc),widths = 15)
writeData(wb,"FINAL Escalation",LeadsNotFound_fo2CEsc,borders = "all",headerStyle = hs1)



o2CEscpath2 <- paste('./Output/',thisDate,"\\o2CEscalation/No Leads found in CRM - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,o2CEscpath2,overwrite = T)


# upload_1 ----------------------------------------------------------------------


Esc_up1 <-Escalation_dump1 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 1",Comment="Escalation Email 1 Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Esc_up2 <-Escalation_dump2 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 2",Comment="Escalation Email 2 Sent to customer",
                                                                
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")
Esc_up3 <-Escalation_dump3 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 3",Comment="Escalation Email 3 Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

fEsc_up3 <-FEscalation %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 3",Comment="Final Escalation Sent to customer",
                                                            `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Esacaltion_up1 <-rbind(Esc_up1,Esc_up2,Esc_up3,fEsc_up3)


if (nrow(Esacaltion_up1)>0) {
  
  NewLeads <- Esacaltion_up1
  FilePath <- paste('./Output/',thisDate,"/Escalation/Escalation Up1- ",nrow(Esacaltion_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}else{
  
  Esacaltion_up1 <- Esacaltion_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- Esacaltion_up1
  FilePath <- paste('./Output/',thisDate,"\\Escalation/Escalation Up1- No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


print("01C_Completed")
##################################

# O2c_upload_1 ----------------------------------------------------------------------


o2CEsc_up1 <-o2CEscalation_dump1 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Escalation Email 1",Comment="02C Escalation Email 1 Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CEsc_up2 <-o2CEscalation_dump2 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Escalation Email 2",Comment="02C Escalation Email 2 Sent to customer",
                                                                      
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")
o2CEsc_up3 <-o2CEscalation_dump3 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Escalation Email 3",Comment="02C Escalation Email 3 Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

fo2CEsc_up3 <-Fo2CEscalation %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Final Escalation Email",Comment="02C Final Escalation Sent to customer",
                                                                  `Allocation Status`="Nil",`Alternate Acc No`="Nil")


o2CEscaltion_up1 <-rbind(o2CEsc_up1,o2CEsc_up2,o2CEsc_up3,fo2CEsc_up3)


if (nrow(o2CEscaltion_up1)>0) {
  
  NewLeads <- o2CEscaltion_up1
  FilePath <- paste('./Output/',thisDate,"/o2CEscalation/o2CEscalation Up1- ",nrow(o2CEscaltion_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}else{
  
  o2CEscaltion_up1 <- o2CEscaltion_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- o2CEscaltion_up1
  FilePath <- paste('./Output/',thisDate,"\\o2CEscalation/o2CEscalation Up1- No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


print("02C_Completed")
