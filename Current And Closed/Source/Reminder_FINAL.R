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

source('.\\Functionfile.R')

today_dat <- Sys.Date()
 

date1 <-format(Sys.Date()-1,"%d%m")

thisDate = Sys.Date()
if(!dir.exists(paste(".\\Output\\",thisDate,sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,"\\Reminder",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\Reminder",sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,"\\o2CReminder",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\o2CReminder",sep=""))
} 




#Read New cases dump --------------------------------


dump_path <- paste(".\\Input\\cis_dump_new_cases.csv",sep='') #%>% filter(prodstatus %in% c('CIS-WP-LP-310','CIS-WP-LP-220'), LDB %in% c('No') & Facilitystatus %in% c('01C Requested customer to send Mail','02C  Customer sent Cut and Paste to Lender'))

dump <- fread(dump_path,na.strings = c('',NA))

dump <- dump[!is.na(dump$email_id),]

dump$CISserialno <- as.numeric(dump$CISserialno)

Rem1_data <-paste("./Input/CUT_AND_PASTE_REMINDER_ONE_",date1,".CSV",sep = '')

Rem1 <-fread(Rem1_data) %>% select('lead_id')
Rem1$lead_id <- as.numeric(Rem1$lead_id)

Rem2_data <-paste("./Input/CUT_AND_PASTE_REMINDER_TWO_",date1,".CSV",sep = '')

Rem2 <-fread(Rem2_data) %>% select('lead_id') 

Rem2$lead_id <- as.numeric(Rem2$lead_id)
Rem3_data <-paste("./Input/CUT_AND_PASTE_REMINDER_THREE_",date1,".CSV",sep = '')

Rem3 <-fread(Rem3_data) %>% select('lead_id') 
Rem3$lead_id <- as.numeric(Rem3$lead_id)


o2CRem1_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_ONE_",date1,".CSV",sep = '')

o2CRem1 <-fread(o2CRem1_data) %>% select('lead_id')
o2CRem1$lead_id <- as.numeric(o2CRem1$lead_id)

o2CRem2_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_TWO_",date1,".CSV",sep = '')

o2CRem2 <-fread(o2CRem2_data) %>% select('lead_id') 

o2CRem2$lead_id <- as.numeric(o2CRem2$lead_id)
o2CRem3_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_THREE_",date1,".CSV",sep = '')

o2CRem3 <-fread(o2CRem3_data) %>% select('lead_id') 
o2CRem3$lead_id <- as.numeric(o2CRem3$lead_id)

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

#reminder1

Reminder_1_file <- dump %>% filter(CISserialno %in% Rem1$lead_id)

Reminder_1_file<-Reminder_1_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_rem1 <-Rem1 %>% filter(!lead_id  %in% Reminder_1_file$CISserialno)

Reminder_1 <- Reminder_1_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender")

o2CReminder_1_file <- dump %>% filter(CISserialno %in% o2CRem1$lead_id)


o2CReminder_1_file<-o2CReminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310")) 

LeadsNotFound_o2Crem1 <-o2CRem1 %>% filter(!lead_id  %in% o2CReminder_1_file$CISserialno)

o2CReminder_1 <- o2CReminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")

Rem1_02C <-o2CReminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender") %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                                                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                                                                            'LDB','Flag')
#reminder2

Reminder_2_file <- dump %>% filter(CISserialno %in% Rem2$lead_id)

Reminder_2_file<-Reminder_2_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_rem2 <-Rem2 %>% filter(!lead_id  %in% Reminder_2_file$CISserialno)

Reminder_2 <- Reminder_2_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

o2CReminder_2_file <- dump %>% filter(CISserialno %in% o2CRem2$lead_id)


o2CReminder_2_file<-o2CReminder_2_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))
LeadsNotFound_o2Crem2 <-o2CRem2 %>% filter(!lead_id  %in% o2CReminder_2_file$CISserialno)

o2CReminder_2 <- o2CReminder_2_file 

# Rem2_02C <-o2CReminder_2_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")%>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                 'LDB','Flag')
#reminder3

Reminder_3_file <- dump %>% filter(CISserialno %in% Rem3$lead_id) 

Reminder_3_file<-Reminder_3_file%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_rem3 <-Rem3 %>% filter(!lead_id  %in% Reminder_3_file$CISserialno)

Reminder_3 <- Reminder_3_file #%>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender")


o2CReminder_3_file <- dump %>% filter(CISserialno %in% o2CRem3$lead_id) 

o2CReminder_3_file<-o2CReminder_3_file%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_o2Crem3 <-o2CRem3 %>% filter(!lead_id  %in% o2CReminder_3_file$CISserialno)

o2CReminder_3 <- o2CReminder_3_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")

# Rem3_02C <-o2CReminder_3_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")%>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                 'LDB','Flag')
# 02C Reminder Email----------------------------------------------------------------
# 
# Reminder_02C <- rbind(Rem1_02C,Rem2_02C,Rem3_02C)
# 
# if (nrow(Reminder_02C)>0) {
#   
#   
#   email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
#   email_list <- read.xlsx(email_fname, sheet ='email')
#   product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
#   
#   dump_02c <- left_join(Reminder_02C, product_family,
#                         by = c('Account_type' = 'account_type')) %>% 
#     mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
#     left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
#     distinct(CISserialno, .keep_all = T)
#   
#   
#   Email_02c_no_mail_id <- dump_02c[is.na(dump_02c$Email),]
#   Email_02c_dump <- dump_02c[!is.na(dump_02c$Email),] %>% distinct(CISserialno, .keep_all = T)
#   
#   NewLeads <- Email_02c_dump
#   FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-02C Email - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
#   
#   
#   lapply(Email_02c_dump$CISserialno, function(lead){
#     
#     #browser()
#     accounts <- Email_02c_dump[Email_02c_dump$CISserialno == lead,]
#     
#     sender <-  'care@creditmantri.com'
#     #recipients <-  accounts$Email_Id
#     bcc_recipients <- 'Ciscare@creditmantri.com'
#     message  <- str_interp('Reg : Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
#     
#     body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>This is regarding the Credit Improvement Services that you had subscribed to earlier, to resolve issues in your credit report with respect to the referenced account.<br/>')
#     
#     body2 <- str_interp('<br/>Thank you for forwarding the mail to the lender. You would receive a response from them in 7 - 10 working days. Requesting you to kindly check your mailbox including junk and spam folders and forward this email to us.<br/>')
#     
#     body3 <- '<br/>Please forward us the response that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/><br />Thanks and Regards <br />Team CreditMantri'
#     
#     email_body <- paste(body1, body2, body3)
#     
#     recipients <- accounts$Email_Id
#     
#     mail_function(sender, recipients, cc_recipients = NULL ,bcc_recipients,message, email_body,filename= NULL)
#     
#     print(paste("Reminder 02C -",accounts$CISserialno))
#     
#   })
#   
# }else{
#   
#   Email_02c_no_mail_id <-Reminder_02C
#   Email_02c_dump <- Reminder_02C
#   
#   NewLeads <- Reminder_02C
#   FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-02C Email_No Leads found - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
# }


#Reminder_1 --------------------------------------------------------------------------


if (nrow(Reminder_1)>0) {
  
  
  Reminder_1 <- Reminder_1 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                      'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                      'LDB','Flag')
  
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Rem_dump1 <- left_join(Reminder_1, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Reminder_no_mail_id_1 <- Rem_dump1[is.na(Rem_dump1$Email),]
  Reminder_dump1 <- Rem_dump1[!is.na(Rem_dump1$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- Reminder_dump1
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
  #Mail function
  
  lapply(Reminder_dump1$CISserialno, function(lead){
    
    #browser()
    accounts <- Reminder_dump1[Reminder_dump1$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : Reminder 1 Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>We had sent you an email a few days ago regarding your subscription to the Credit Improvement Service with CreditMantri. Since we have not heard from you, we are following-up to remind you. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com <br/> <br/><b>Kindly ignore this mail if you have already sent the below text to the lender.</b><br/><br/>Please forward us any responses that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/>')
    
    body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    #recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    print(paste("Reminder one -",accounts$CISserialno))
    
  })
  
  
  
}else{
  Reminder_1 <- Reminder_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  Reminder_no_mail_id_1 <-Reminder_1
  Reminder_dump1 <- Reminder_1
  
  NewLeads <- Reminder_1
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
}



#Reminder_2 --------------------------------------------------------------------------

if (nrow(Reminder_2)>0) {
  
  
  Reminder_2 <- Reminder_2 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                      'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                      'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Rem_dump2 <- left_join(Reminder_2, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Reminder_no_mail_id_2 <- Rem_dump2[is.na(Rem_dump2$Email),]
  Reminder_dump2 <- Rem_dump2[!is.na(Rem_dump2$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- Reminder_dump2
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  #mail function
  
  lapply(Reminder_dump2$CISserialno, function(lead){
    
    #browser()
    accounts <- Reminder_dump2[Reminder_dump2$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : Reminder 2 Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>This is our  <b>2nd reminder </b>regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from you regarding sending out the email to your lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com<br/><br/><b>Kindly ignore this mail if you have already sent the below text to the lender.</b><br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution. <br/>')
    
    body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    print(paste("Reminder two -",accounts$CISserialno))
    
  })
  
  
  
  
}else{
  
  Reminder_2 <- Reminder_2 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  Reminder_no_mail_id_2 <-Reminder_2
  Reminder_dump2 <- Reminder_2
  
  NewLeads <- Reminder_2
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}

#Reminder_3 --------------------------------------------------------------------------



if (nrow(Reminder_3)>0) {
  
  
  
  Reminder_3 <- Reminder_3 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                      'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                      'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  Rem_dump3 <- left_join(Reminder_3, product_family,
                         by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  Reminder_no_mail_id_3 <- Rem_dump3[is.na(Rem_dump3$Email),]
  Reminder_dump3 <- Rem_dump3[!is.na(Rem_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- Reminder_dump3
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  #Mail function
  
  lapply(Reminder_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- Reminder_dump3[Reminder_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : Reminder 3 Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>This is our <b>3rd reminder </b>regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from you regarding sending out the email to your lender and hence this reminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com.<br/> <br/><b>Kindly ignore this mail if you have already sent the below given text to the bank.</b><br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.<br/>')
    
    body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    
    print(paste("Reminder three -",accounts$CISserialno))
  })
  
  
}else{
  
  Reminder_3 <- Reminder_3 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  Reminder_no_mail_id_3 <-Reminder_3
  Reminder_dump3 <- Reminder_3
  
  NewLeads <- Reminder_3
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}





#o2CReminder_1 --------------------------------------------------------------------------


if (nrow(o2CReminder_1)>0) {
  
  
  o2CReminder_1 <- o2CReminder_1 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2Creminder-1.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CRem_dump1 <- left_join(o2CReminder_1, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CReminder_no_mail_id_1 <- o2CRem_dump1[is.na(o2CRem_dump1$Email),]
  o2CReminder_dump1 <- o2CRem_dump1[!is.na(o2CRem_dump1$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- o2CReminder_dump1
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
  #Mail function
  
  lapply(o2CReminder_dump1$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CReminder_dump1[o2CReminder_dump1$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : 02CReminder 1- ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>We had sent you an email a few days ago regarding your subscription to the Credit Improvement Service with CreditMantri. Since we have not heard from you, we are following-up to remind you. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com <br/><br/>Please forward us any responses that you have received, or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/>')
    
    body2 <- str_interp('<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  I had sent an email a few days ago that on a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible. <br/><br/>Thanks and Regards <br/>${accounts$Customer_Name}')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    print(paste("o2CReminder one -",accounts$CISserialno))
    
  })
  
  
  
}else{
  o2CReminder_1 <- o2CReminder_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_1 <-o2CReminder_1
  o2CReminder_dump1 <- o2CReminder_1
  
  NewLeads <- o2CReminder_1
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
}

#o2CReminder_2 --------------------------------------------------------------------------

if (nrow(o2CReminder_2)>0) {
  
  
  o2CReminder_2 <- o2CReminder_2 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2Creminder-2.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CRem_dump2 <- left_join(o2CReminder_2, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CReminder_no_mail_id_2 <- o2CRem_dump2[is.na(o2CRem_dump2$Email),]
  o2CReminder_dump2 <- o2CRem_dump2[!is.na(o2CRem_dump2$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- o2CReminder_dump2
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  #mail function
  
  lapply(o2CReminder_dump2$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CReminder_dump2[o2CReminder_dump2$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : 02CReminder 2- ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>This is our  <b>2nd o2CReminder </b>regarding the Credit Improvement Service that you had subscribed to with CreditMantri and we have sent to lender. We have not yet received confirmation from lender regarding sending out the email to your lender and hence this o2CReminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We request you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com <br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.')
    
    body2 <- str_interp('<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  I had sent two emails a few days ago and since I have not heard from you, I am following up to remind you that on scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible. <br/><br/>Thanks and Regards <br/>${accounts$Customer_Name}')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    print(paste("o2CReminder two -",accounts$CISserialno))
    
  })
  
  
  
  
}else{
  
  o2CReminder_2 <- o2CReminder_2 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_2 <-o2CReminder_2
  o2CReminder_dump2 <- o2CReminder_2
  
  NewLeads <- o2CReminder_2
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}



#o2CReminder_3 --------------------------------------------------------------------------



if (nrow(o2CReminder_3)>0) {
  
  
  
  o2CReminder_3 <- o2CReminder_3 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2Creminder-3.xlsx'
  email_list <- read.xlsx(email_fname, sheet ='email')
  product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
  
  o2CRem_dump3 <- left_join(o2CReminder_3, product_family,
                            by = c('Account_type' = 'account_type')) %>% 
    mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
    left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
    distinct(CISserialno, .keep_all = T)
  
  
  o2CReminder_no_mail_id_3 <- o2CRem_dump3[is.na(o2CRem_dump3$Email),]
  o2CReminder_dump3 <- o2CRem_dump3[!is.na(o2CRem_dump3$Email),] %>% distinct(CISserialno, .keep_all = T)
  
  
  NewLeads <- o2CReminder_dump3
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  #Mail function
  
  lapply(o2CReminder_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- o2CReminder_dump3[o2CReminder_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : 02CReminder 3- ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>This is our <b>3rd o2CReminder </b>regarding the Credit Improvement Service that you had subscribed to with CreditMantri. We have not yet received confirmation from lender and hence this o2CReminder. We had contacted your lender to seek the amount due on your account. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address to furnish the details on the account. We had requested you to cut and paste the below provided text and send it to the lender ( ${accounts$Email} ) from your registered email address. Kindly retain the subject line. Please copy us as a cc in your email - care@creditmantri.com. <br/><br/>We request you to forward to us any responses that you have received, or you may receive from the lender to care@creditmantri.com to further proceed with your account resolution.')
    
    body2 <- str_interp('<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Ref: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>  This is my third o2CReminder about the email that I had sent earlier too and since I have not heard from you, I am following up to remind you that on scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a 30-59 days past due status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible. <br/>I have not received any response to my previous mails to you. Request you to please help with urgent closure. <br/><br/>Thanks and Regards <br/>${accounts$Customer_Name}')
    
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br/>Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body,filename= NULL)
    
    
    print(paste("o2CReminder three -",accounts$CISserialno))
  })
  
  
}else{
  
  o2CReminder_3 <- o2CReminder_3 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_3 <-o2CReminder_3
  o2CReminder_dump3 <- o2CReminder_3
  
  NewLeads <- o2CReminder_3
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}





#unsent Email ----------------------------------------------------------------------------------------

wb <- createWorkbook()
addWorksheet(wb,"Reminder-1")
addWorksheet(wb,"Reminder-2")
addWorksheet(wb,"Reminder-3")
addWorksheet(wb,"o2CReminder-1")
addWorksheet(wb,"o2CReminder-2")
addWorksheet(wb,"o2CReminder-3")

hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"Reminder-1",cols = 1:ncol(Reminder_no_mail_id_1),widths = 15)
writeData(wb,"Reminder-1",Reminder_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-2",cols = 1:ncol(Reminder_no_mail_id_2),widths = 15)
writeData(wb,"Reminder-2",Reminder_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-3",cols = 1:ncol(Reminder_no_mail_id_3),widths = 15)
writeData(wb,"Reminder-3",Reminder_no_mail_id_3,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-1",cols = 1:ncol(o2CReminder_no_mail_id_1),widths = 15)
writeData(wb,"o2CReminder-1",o2CReminder_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-2",cols = 1:ncol(o2CReminder_no_mail_id_2),widths = 15)
writeData(wb,"o2CReminder-2",o2CReminder_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-3",cols = 1:ncol(o2CReminder_no_mail_id_3),widths = 15)
writeData(wb,"o2CReminder-3",o2CReminder_no_mail_id_3,borders = "all",headerStyle = hs1)


path <- paste('./Output/',thisDate,"\\Reminder/Reminder Unsent Emails on - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)

path2 <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder Unsent Emails on - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path2,overwrite = T)




#No leads found in Crm---------------------------------------------------------------


wb <- createWorkbook()

addWorksheet(wb,"Reminder-1")
addWorksheet(wb,"Reminder-2")
addWorksheet(wb,"Reminder-3")
addWorksheet(wb,"o2CReminder-1")
addWorksheet(wb,"o2CReminder-2")
addWorksheet(wb,"o2CReminder-3")


hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")



setColWidths(wb,"Reminder-1",cols = 1:ncol(LeadsNotFound_rem1),widths = 15)
writeData(wb,"Reminder-1",LeadsNotFound_rem1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-2",cols = 1:ncol(LeadsNotFound_rem2),widths = 15)
writeData(wb,"Reminder-2",LeadsNotFound_rem2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-3",cols = 1:ncol(LeadsNotFound_rem3),widths = 15)
writeData(wb,"Reminder-3",LeadsNotFound_rem3,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-1",cols = 1:ncol(LeadsNotFound_o2Crem1),widths = 15)
writeData(wb,"o2CReminder-1",LeadsNotFound_o2Crem1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-2",cols = 1:ncol(LeadsNotFound_o2Crem2),widths = 15)
writeData(wb,"o2CReminder-2",LeadsNotFound_o2Crem2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-3",cols = 1:ncol(LeadsNotFound_o2Crem3),widths = 15)
writeData(wb,"o2CReminder-3",LeadsNotFound_o2Crem3,borders = "all",headerStyle = hs1)


path <- paste('./Output/',thisDate,"\\Reminder/Reminder No Leads found in CRM - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)

path2 <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder No Leads found in CRM - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path2,overwrite = T)


# upload_1 ----------------------------------------------------------------------

Remi_up1 <-Reminder_dump1  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder one",Comment="Cut and Paste Reminder one Sent to customer",
                                                                   `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Remi_up2 <-Reminder_dump2  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder two",Comment="Cut and Paste Reminder two Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Remi_up3 <-Reminder_dump3  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder three",Comment="Cut and Paste Reminder three Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

# Rem_02C_up1 <-Email_02c_dump %>% select('CISserialno') %>% mutate(`Facility Status`="Nil",`CRM Status`="Nil",Comment="02C Reminder Email sent customer",
#                                                                   `Allocation Status`="Nil",`Alternate Acc No`="Nil")
# 
Reminder_up1 <-rbind(Remi_up1,Remi_up2,Remi_up3)

if (nrow(Reminder_up1)>0) {
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder Up1- ",nrow(Reminder_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}else{
  
  Reminder_up1 <- Reminder_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"\\Reminder/Reminder Up1-No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}

print("01C_Completed")



o2CRemi_up1 <-o2CReminder_dump1  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Cut and Paste Reminder one",Comment="02C Cut and Paste Reminder one Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CRemi_up2 <-o2CReminder_dump2  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Cut and Paste Reminder two",Comment="02C Cut and Paste Reminder two Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CRemi_up3 <-o2CReminder_dump3  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="02C Cut and Paste Reminder three",Comment="02C Cut and Paste Reminder three Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")
o2CReminder_up1 <-rbind(o2CRemi_up1,o2CRemi_up2,o2CRemi_up3)


if (nrow(o2CReminder_up1)>0) {
  
  NewLeads <- o2CReminder_up1
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder Up1- ",nrow(o2CReminder_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}else{
  
  o2CReminder_up1 <- o2CReminder_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- o2CReminder_up1
  FilePath <- paste('./Output/',thisDate,"\\o2CReminder/o2CReminder Up1-No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


print("02C_Completed")
























