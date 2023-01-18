rm(list = ls())
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_172')
#install.packages("formattable")  
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

if(!dir.exists(paste(".\\Output\\",thisDate,"\\Upload files",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\Upload files",sep=""))
} 

#Read New cases dump --------------------------------


dump_path <- paste(".\\Input\\cis_dump_new_cases.csv",sep='')

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
Esc1_data <-paste("./Input/ESCALATION_EMAIL_1_",date1,".CSV",sep = '')

Esc1 <-fread(Esc1_data) %>% select('lead_id') 
Esc1$lead_id <- as.numeric(Esc1$lead_id)
Esc2_data <-paste("./Input/ESCALATION_EMAIL_2_",date1,".CSV",sep = '')

Esc2 <-fread(Esc2_data) %>% select('lead_id')
Esc2$lead_id <- as.numeric(Esc2$lead_id)
Esc3_data <-paste("./Input/ESCALATION_EMAIL_3_",date1,".CSV",sep = '')

Esc3 <-fread(Esc3_data) %>% select('lead_id') 
Esc3$lead_id <- as.numeric(Esc3$lead_id)



fEsc_data <-paste("./Input/FINAL_ESCALATION_",date1,".CSV",sep = '')

fEsc <-fread(fEsc_data) %>% select('lead_id') 
fEsc$lead_id <-as.numeric(fEsc$lead_id)


fo2CEsc_data <-paste("./Input/02C_FINAL_ESCALATION_",date1,".CSV",sep = '')

fo2CEsc <-fread(fo2CEsc_data) %>% select('lead_id') 
fo2CEsc$lead_id <-as.numeric(fo2CEsc$lead_id)



o2CRem1_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_ONE_",date1,".CSV",sep = '')

o2CRem1 <-fread(o2CRem1_data) %>% select('lead_id')
o2CRem1$lead_id <- as.numeric(o2CRem1$lead_id)

o2CRem2_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_TWO_",date1,".CSV",sep = '')

o2CRem2 <-fread(o2CRem2_data) %>% select('lead_id') 

o2CRem2$lead_id <- as.numeric(o2CRem2$lead_id)
o2CRem3_data <-paste("./Input/02C_CUT_AND_PASTE_REMINDER_THREE_",date1,".CSV",sep = '')

o2CRem3 <-fread(o2CRem3_data) %>% select('lead_id') 
o2CRem3$lead_id <- as.numeric(o2CRem3$lead_id)
o2CEsc1_data <-paste("./Input/02C_ESCALATION_EMAIL_1_",date1,".CSV",sep = '')

o2CEsc1 <-fread(o2CEsc1_data) %>% select('lead_id') 
o2CEsc1$lead_id <- as.numeric(o2CEsc1$lead_id)
o2CEsc2_data <-paste("./Input/02C_ESCALATION_EMAIL_2_",date1,".CSV",sep = '')

o2CEsc2 <-fread(o2CEsc2_data) %>% select('lead_id')
o2CEsc2$lead_id <- as.numeric(o2CEsc2$lead_id)
o2CEsc3_data <-paste("./Input/02C_ESCALATION_EMAIL_3_",date1,".CSV",sep = '')

o2CEsc3 <-fread(o2CEsc3_data) %>% select('lead_id') 
o2CEsc3$lead_id <- as.numeric(o2CEsc3$lead_id)

cut_and_paste <- dump %>% filter(LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310") & Facilitystatus == "00 To Start Action")





#cut and paste format ------------------------------


if (nrow(cut_and_paste)>0) {

cut_and_paste <- cut_and_paste %>% 
  mutate(
    Account_status = case_when(
    
    (account_status == "Closed Account" & asset_classification == "Speci") ~ "SPECIAL MENTION",
    (account_status == "Closed Account" & asset_classification == "Sub-s") ~ "SUB-STANDARD",
    (account_status == "Closed Account" & asset_classification == "Doubt") ~ "DOUBTFUL",
    (account_status == "Current Account" & asset_classification == "Speci") ~ "SPECIAL MENTION",
    (account_status == "Current Account" & asset_classification == "Sub-s") ~ "SUB-STANDARD",
    (account_status == "Current Account" & asset_classification == "Doubt") ~ "DOUBTFUL",
    
    T ~ account_status
  
))

cut_and_paste <- cut_and_paste %>% filter(Account_status != "Closed Account" & Account_status != "Current Account")


cut_and_paste$Customer_Name=paste(cut_and_paste$first_name,cut_and_paste$last_name)
cut_and_paste$Date=cut_and_paste$subscription_date
cut_and_paste$Phone_Number=cut_and_paste$phone_home
cut_and_paste$Address=cut_and_paste$city
cut_and_paste$Product_Status=cut_and_paste$prodstatus
cut_and_paste <- cut_and_paste %>% mutate(Flag="")
cut_and_paste$Lender_Name=cut_and_paste$lender_name
cut_and_paste$Account_type=cut_and_paste$account_type
cut_and_paste$Email_Id=cut_and_paste$email_id

cut_and_paste <- cut_and_paste %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                          'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
cut_and_paste$CISserialno <- as.numeric(cut_and_paste$CISserialno)

#file creation

NewLeads <- cut_and_paste
FilePath <- path <- paste('./Output/',thisDate,"/Upload files/Cut and paste - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)



email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
email_list <- read.xlsx(email_fname, sheet ='email')
product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')

dump1 <- left_join(cut_and_paste, product_family,
                   by = c('Account_type' = 'account_type')) %>% 
  mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
  left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
  distinct(CISserialno, .keep_all = T)


Cut_and_paste_no_mail_id <- dump1[is.na(dump1$Email),]
Cut_and_paste_dump1 <- dump1[!is.na(dump1$Email),] %>% distinct(CISserialno, .keep_all = T)


}else{
  cut_and_paste <- cut_and_paste %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  cut_and_paste$CISserialno <- as.numeric(cut_and_paste$CISserialno)
  NewLeads <- cut_and_paste
  FilePath <- paste('./Output/',thisDate,"/Upload files/Cut and paste - No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  }

#02C


#Current and closed -------------------------------------------

Current_and_closed <-dump %>% filter(LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310") & Facilitystatus == "00 To Start Action") 
  
Current_and_closed <- Current_and_closed %>% filter(account_status %in% c("Current Account","Closed Account"))

if (nrow(Current_and_closed)>0) {
  


Current_and_closed$Customer_Name=paste(Current_and_closed$first_name,Current_and_closed$last_name)
Current_and_closed$Date=Current_and_closed$subscription_date
Current_and_closed$Phone_Number=Current_and_closed$phone_home
Current_and_closed$Address=Current_and_closed$city
Current_and_closed$Product_Status=Current_and_closed$prodstatus
Current_and_closed$Flag=""


Current_and_closed <- Current_and_closed %>% select('CISserialno','Date','Customer_Name','lender_name','account_type','Account_No',
                                          'account_status','Facilitystatus','email_id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
Current_and_closed$CISserialno <- as.numeric(Current_and_closed$CISserialno)

NewLeads <- Current_and_closed
FilePath <- paste('./Output/',thisDate,"/Upload files/Current and closed - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)


#mail function


  sender <-  'ops-cis@creditmantri.com'
  #recipients <-  accounts$Email_Id
  cc_recipients <- 'ops-cis@creditmantri.com'
  message  <- str_interp('Current And Closed Account for - ${Sys.Date()}' )
  
  body1 <- str_interp('<br/>Dear Team,<br/>')
  
  body2 <- str_interp('<br/> Please find the attached Current and closed for - ${Sys.Date()}')
  
  body3 <- '<br/><br/>Thanks and Regards <br />OPS Team'
  
  email_body <- paste(body1, body2, body3)
  
  recipients <- c('azharudeen@creditmantri.com')
  
  filename <- paste('./Output/',thisDate,'/Upload files/Current and closed - ',thisDate,'.xlsx',sep = '')
  
  mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body, filename)
  
  
}else{
 
  sender <-  'ops-cis@creditmantri.com'
  #recipients <-  accounts$Email_Id
  cc_recipients <- 'ops-cis@creditmantri.com'
  message  <- str_interp('Current And Closed Account for - ${Sys.Date()}' )
  
  body1 <- str_interp('<br/>Dear Team,<br/>')
  
  body2 <- str_interp('<br/> No Leads found for Current and closed on - ${Sys.Date()}')
  
  body3 <- '<br/><br/>Thanks and Regards <br />OPS Team'
  
  email_body <- paste(body1, body2, body3)
  
  recipients <- c('azharudeen@creditmantri.com','ops-cis@creditmantri.com')
  
  #filename <- paste('./Output/',thisDate,'/Current and closed - ',thisDate,'.xlsx',sep = '')
  
  mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body, filename=NULL)
  
}


# ######################02C
# 
# 
# o2Ccut_and_paste <- dump %>% filter(prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310") & Facilitystatus == "01C Requested customer to send Mail")
# 
# 
# 
# 
# 
# #cut and paste format ------------------------------
# 
# 
# if (nrow(o2Ccut_and_paste)>0) {
#   
#   o2Ccut_and_paste <- o2Ccut_and_paste %>% 
#     mutate(
#       Account_status = case_when(
#         
#         (account_status == "Closed Account" & asset_classification == "Speci") ~ "SPECIAL MENTION",
#         (account_status == "Closed Account" & asset_classification == "Sub-s") ~ "SUB-STANDARD",
#         (account_status == "Closed Account" & asset_classification == "Doubt") ~ "DOUBTFUL",
#         (account_status == "Current Account" & asset_classification == "Speci") ~ "SPECIAL MENTION",
#         (account_status == "Current Account" & asset_classification == "Sub-s") ~ "SUB-STANDARD",
#         (account_status == "Current Account" & asset_classification == "Doubt") ~ "DOUBTFUL",
#         
#         T ~ account_status
#         
#       ))
#   
#   o2Ccut_and_paste <- o2Ccut_and_paste %>% filter(Account_status != "Closed Account" & Account_status != "Current Account")
#   
#   
#   o2Ccut_and_paste$Customer_Name=paste(o2Ccut_and_paste$first_name,o2Ccut_and_paste$last_name)
#   o2Ccut_and_paste$Date=o2Ccut_and_paste$subscription_date
#   o2Ccut_and_paste$Phone_Number=o2Ccut_and_paste$phone_home
#   o2Ccut_and_paste$Address=o2Ccut_and_paste$city
#   o2Ccut_and_paste$Product_Status=o2Ccut_and_paste$prodstatus
#   o2Ccut_and_paste <- o2Ccut_and_paste %>% mutate(Flag="")
#   o2Ccut_and_paste$Lender_Name=o2Ccut_and_paste$lender_name
#   o2Ccut_and_paste$Account_type=o2Ccut_and_paste$account_type
#   o2Ccut_and_paste$Email_Id=o2Ccut_and_paste$email_id
#   
#   o2Ccut_and_paste <- o2Ccut_and_paste %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                   'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                   'LDB','Flag')
#   o2Ccut_and_paste$CISserialno <- as.numeric(o2Ccut_and_paste$CISserialno)
#   
#   #file creation
#   
#   NewLeads <- o2Ccut_and_paste
#   FilePath <- path <- paste('./Output/',thisDate,"/Upload files/o2CCut and paste - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
#   
#   
#   
#   email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-o2CCut and paste & reminder.xlsx'
#   email_list <- read.xlsx(email_fname, sheet ='email')
#   product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
#   
#   dump1 <- left_join(o2Ccut_and_paste, product_family,
#                      by = c('Account_type' = 'account_type')) %>% 
#     mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
#     left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
#     distinct(CISserialno, .keep_all = T)
#   
#   
#   o2Ccut_and_paste_no_mail_id <- dump1[is.na(dump1$Email),]
#   o2Ccut_and_paste_dump1 <- dump1[!is.na(dump1$Email),] %>% distinct(CISserialno, .keep_all = T)
#   
#   
# }else{
#   o2Ccut_and_paste <- o2Ccut_and_paste %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                  'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                  'LDB','Flag')
#   
#   o2Ccut_and_paste$CISserialno <- as.numeric(o2Ccut_and_paste$CISserialno)
#   NewLeads <- o2Ccut_and_paste
#   FilePath <- paste('./Output/',thisDate,"/Upload files/o2CCut and paste - No Leads found - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
#   
#   
# }

#online request--------------------------------------------------------------------



Online <- cut_and_paste %>% filter(Lender_Name == 'Allahabad')


if (nrow(Online)>0) {

NewLeads <- Online
FilePath <- paste('./Output/',thisDate,"/Upload files/Online Request - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)

#mail function

sender <-  'ops-cis@creditmantri.com'

cc_recipients <- NULL
message  <- str_interp('Online Request for - ${Sys.Date()}' )

body1 <- str_interp('<br/>Dear Team,<br/>')

body2 <- str_interp('<br/> Please find the attached Online Request for for - ${Sys.Date()}')

body3 <- '<br/><br/>Thanks and Regards <br />OPS Team'

email_body <- paste(body1, body2, body3)

recipients <- 'ops-cis@creditmantri.com'

filename <- paste('./Output/',thisDate,'/Upload files/Online Request - ',thisDate,'.xlsx',sep = '')

mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body, filename)


}else{
  
  sender <-  'ops-cis@creditmantri.com
'
  
  cc_recipients <- 'ops-cis@creditmantri.com'
  message  <- str_interp('Online Request for - ${Sys.Date()}' )
  
  body1 <- str_interp('<br/>Dear Team,<br/>')
  
  body2 <- str_interp('<br/> No leads found for Online Request on - ${Sys.Date()}')
  
  body3 <- '<br/><br/>Thanks and Regards <br />OPS Team'
  
  email_body <- paste(body1, body2, body3)
  
  recipients <- 'ops-cis@creditmantri.com'
  
  #filename <- paste('./Output/',thisDate,'/Online Request - ',thisDate,'.xlsx',sep = '')
  
  mail_function(sender, recipients, cc_recipients,bcc_recipients = NULL ,message, email_body, filename=NULL)
  
}


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



#Reminder_1 --------------------------------------------------------------------------


 

Reminder_1_file <- dump %>% filter(CISserialno %in% Rem1$lead_id)

Reminder_1_file<-Reminder_1_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_rem1 <-Rem1 %>% filter(!lead_id  %in% Reminder_1_file$CISserialno)

Reminder_1 <- Reminder_1_file 

# Rem1_02C <-Reminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310")) %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                                                                                    'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status','LDB','Flag')
#  
# 
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
FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-1 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)



}else{
  Reminder_1 <- Reminder_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  Reminder_no_mail_id_1 <-Reminder_1
  Reminder_dump1 <- Reminder_1
  
  NewLeads <- Reminder_1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
}



#Reminder_2 --------------------------------------------------------------------------



Reminder_2_file <- dump %>% filter(CISserialno %in% Rem2$lead_id)

LeadsNotFound_rem2 <-Rem2 %>% filter(!lead_id  %in% Reminder_2_file$CISserialno)

Reminder_2 <- Reminder_2_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

# Rem2_02C <-Reminder_2_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310")) %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status','LDB','Flag')
# 
# 
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
FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-2 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)

  

}else{
  
  Reminder_2 <- Reminder_2 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  Reminder_no_mail_id_2 <-Reminder_2
  Reminder_dump2 <- Reminder_2

  NewLeads <- Reminder_2
  FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}

#Reminder_3 --------------------------------------------------------------------------



Reminder_3_file <- dump %>% filter(CISserialno %in% Rem3$lead_id)

LeadsNotFound_rem3 <-Rem3 %>% filter(!lead_id  %in% Reminder_3_file$CISserialno)

Reminder_3 <- Reminder_3_file %>% filter(Facilitystatus != "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))


# Rem3_02C <-Reminder_3_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")%>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                 'LDB','Flag')
# 
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
FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-3 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)


}else{
  
  Reminder_3 <- Reminder_3 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                         'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                         'LDB','Flag')
  
  Reminder_no_mail_id_3 <-Reminder_3
  Reminder_dump3 <- Reminder_3
  
  NewLeads <- Reminder_3
  FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}

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
FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-1 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)



}else{
  
  
  Escalation_1 <- Escalation_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
           'LDB','Flag')
  
  Escalation_no_mail_id_1 <-Escalation_1
  Escalation_dump1 <-Escalation_1

  NewLeads <- Escalation_1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
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
FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-2 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)



}else{
  

  
  Escalation_2 <- Escalation_2 %>% mutate(Flag='')%>%  
    select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                          'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                          'LDB','Flag')
  Escalation_no_mail_id_2 <-Escalation_2
  Escalation_dump2 <-Escalation_2
 
  NewLeads <- Escalation_2
  FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}


#Escalation_3 ----------------------------------------------------------------------



Escalation_3 <- dump %>% filter(CISserialno %in% Esc3$lead_id)

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
FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-3 Email - ",thisDate,'.xlsx',sep='')
FileCreate <- FileCreation(NewLeads,FilePath)



}else{
  

  
  Escalation_3 <- Escalation_3 %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
           'LDB','Flag')
  
  Escalation_no_mail_id_3 <-Escalation_3
  Escalation_dump3 <-Escalation_3

  NewLeads <- Escalation_3
  FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}




#####################02C

#o2CReminder_1 --------------------------------------------------------------------------




o2CReminder_1_file <- dump %>% filter(CISserialno %in% o2CRem1$lead_id)

o2CReminder_1_file<-o2CReminder_1_file%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))


LeadsNotFound_o2CRem1 <-o2CRem1 %>% filter(!lead_id  %in% o2CReminder_1_file$CISserialno)

o2CReminder_1 <- o2CReminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")

# o2CRem1_02C <-o2CReminder_1_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender") %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                        'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                        'LDB','Flag')
# 

if (nrow(o2CReminder_1)>0) {
  
  
  o2CReminder_1 <- o2CReminder_1 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
}else{
  o2CReminder_1 <- o2CReminder_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_1 <-o2CReminder_1
  o2CReminder_dump1 <- o2CReminder_1
  
  NewLeads <- o2CReminder_1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
}



#o2CReminder_2 --------------------------------------------------------------------------



o2CReminder_2_file <- dump %>% filter(CISserialno %in% o2CRem2$lead_id)

o2CReminder_2_file<-o2CReminder_2_file%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))

LeadsNotFound_o2CRem2 <-o2CRem2 %>% filter(!lead_id  %in% o2CReminder_2_file$CISserialno)

o2CReminder_2 <- o2CReminder_2_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")

# o2CRem2_02C <-o2CReminder_2_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")%>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                       'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                       'LDB','Flag')
# 

if (nrow(o2CReminder_2)>0) {
  
  
  o2CReminder_2 <- o2CReminder_2 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
}else{
  
  o2CReminder_2 <- o2CReminder_2 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_2 <-o2CReminder_2
  o2CReminder_dump2 <- o2CReminder_2
  
  NewLeads <- o2CReminder_2
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}

#o2CReminder_3 --------------------------------------------------------------------------



o2CReminder_3_file <- dump %>% filter(CISserialno %in% o2CRem3$lead_id)

o2CReminder_3_file<-o2CReminder_3_file%>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender", LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310"))


LeadsNotFound_o2CRem3 <-o2CRem3 %>% filter(!lead_id  %in% o2CReminder_3_file$CISserialno)

o2CReminder_3 <- o2CReminder_3_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")

# o2CRem3_02C <-o2CReminder_3_file %>% filter(Facilitystatus == "02C  Customer sent Cut and Paste to Lender")%>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
#                                                                                                                       'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
#                                                                                                                       'LDB','Flag')

if (nrow(o2CReminder_3)>0) {
  
  
  
  o2CReminder_3 <- o2CReminder_3 %>% select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                            'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                            'LDB','Flag')
  
  email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}else{
  
  o2CReminder_3 <- o2CReminder_3 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                           'LDB','Flag')
  
  o2CReminder_no_mail_id_3 <-o2CReminder_3
  o2CReminder_dump3 <- o2CReminder_3
  
  NewLeads <- o2CReminder_3
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-1 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
}else{
  
  
  o2CEscalation_1 <- o2CEscalation_1 %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                               'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                               'LDB','Flag')
  
  o2CEscalation_no_mail_id_1 <-o2CEscalation_1
  o2CEscalation_dump1 <-o2CEscalation_1
  
  NewLeads <- o2CEscalation_1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-1 Email_No Leads found - ",thisDate,'.xlsx',sep='')
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-2 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
}else{
  
  
  
  o2CEscalation_2 <- o2CEscalation_2 %>% mutate(Flag='')%>%  
    select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
           'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
           'LDB','Flag')
  o2CEscalation_no_mail_id_2 <-o2CEscalation_2
  o2CEscalation_dump2 <-o2CEscalation_2
  
  NewLeads <- o2CEscalation_2
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-2 Email_No Leads found - ",thisDate,'.xlsx',sep='')
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
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-3 Email - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  
}else{
  
  
  
  o2CEscalation_3 <- o2CEscalation_3 %>%  select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                                 'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                                 'LDB','Flag')
  
  o2CEscalation_no_mail_id_3 <-o2CEscalation_3
  o2CEscalation_dump3 <-o2CEscalation_3
  
  NewLeads <- o2CEscalation_3
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation-3 Email_No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


# 02C Reminder Email----------------------------------------------------------------

#Reminder_02C <- rbind(Rem1_02C,Rem2_02C,Rem3_02C)


# o2CEscalation_2 <- dump %>% filter(CISserialno %in% o2CEsc2$lead_id)
# 
# LeadsNotFound_o2CEsc2 <-o2CEsc2 %>% filter(!lead_id  %in% o2CEscalation_2$CISserialno)
# 
# if (nrow(Reminder_02C)>0) {
#   
#   
#   email_fname <- '.\\Input\\Email IDS\\Lender_Email_Details-cut and paste & reminder.xlsx'
#   email_list <- read.xlsx(email_fname, sheet ='email')
#   product_family <- read.xlsx(email_fname, sheet = 'ProductFamily')
#   
#   dump_02c <- left_join(Reminder_02C, product_family,
#                          by = c('Account_type' = 'account_type')) %>% 
#     mutate(`bank name and product short code` = paste(Lender_Name,product_family_short)) %>% 
#     left_join(email_list, by = c('bank name and product short code' = 'lender.name')) %>% 
#     distinct(CISserialno, .keep_all = T)
#   
#   
#  Email_02c_no_mail_id <- dump_02c[is.na(dump_02c$Email),]
#   Email_02c_dump <- dump_02c[!is.na(dump_02c$Email),] %>% distinct(CISserialno, .keep_all = T)
#   
#   NewLeads <- Email_02c_dump
#   FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-02C Email - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
#   
#   
#   
# }else{
# 
#   Email_02c_no_mail_id <-Reminder_02C
#   Email_02c_dump <- Reminder_02C
# 
#   NewLeads <- Reminder_02C
#   FilePath <- paste('./Output/',thisDate,"/Upload files/Reminder-02C Email_No Leads found - ",thisDate,'.xlsx',sep='')
#   FileCreate <- FileCreation(NewLeads,FilePath)
# }







#unsent Email ----------------------------------------------------------------------------------------

wb <- createWorkbook()
addWorksheet(wb,"Cut and paste")
addWorksheet(wb,"Reminder-1")
addWorksheet(wb,"Reminder-2")
addWorksheet(wb,"Reminder-3")
addWorksheet(wb,"Escalation-1")
addWorksheet(wb,"Escalation-2")
addWorksheet(wb,"Escalation-3")
addWorksheet(wb,"02C Reminder Email")

addWorksheet(wb,"o2CReminder-1")
addWorksheet(wb,"o2CReminder-2")
addWorksheet(wb,"o2CReminder-3")
addWorksheet(wb,"o2CEscalation-1")
addWorksheet(wb,"o2CEscalation-2")
addWorksheet(wb,"o2CEscalation-3")


hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"Cut and paste",cols = 1:ncol(Cut_and_paste_no_mail_id),widths = 15)
writeData(wb,"Cut and paste",Cut_and_paste_no_mail_id,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-1",cols = 1:ncol(Reminder_no_mail_id_1),widths = 15)
writeData(wb,"Reminder-1",Reminder_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-2",cols = 1:ncol(Reminder_no_mail_id_2),widths = 15)
writeData(wb,"Reminder-2",Reminder_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Reminder-3",cols = 1:ncol(Reminder_no_mail_id_3),widths = 15)
writeData(wb,"Reminder-3",Reminder_no_mail_id_3,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-1",cols = 1:ncol(Escalation_no_mail_id_1),widths = 15)
writeData(wb,"Escalation-1",Escalation_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-2",cols = 1:ncol(Escalation_no_mail_id_2),widths = 15)
writeData(wb,"Escalation-2",Escalation_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-3",cols = 1:ncol(Escalation_no_mail_id_3),widths = 15)
writeData(wb,"Escalation-3",Escalation_no_mail_id_3,borders = "all",headerStyle = hs1)

setColWidths(wb,"02C Reminder Email",cols = 1:ncol(Email_02c_no_mail_id),widths = 15)
writeData(wb,"02C Reminder Email",Email_02c_no_mail_id,borders = "all",headerStyle = hs1)


setColWidths(wb,"o2CReminder-1",cols = 1:ncol(o2CReminder_no_mail_id_1),widths = 15)
writeData(wb,"o2CReminder-1",o2CReminder_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-2",cols = 1:ncol(o2CReminder_no_mail_id_2),widths = 15)
writeData(wb,"o2CReminder-2",o2CReminder_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-3",cols = 1:ncol(o2CReminder_no_mail_id_3),widths = 15)
writeData(wb,"o2CReminder-3",o2CReminder_no_mail_id_3,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-1",cols = 1:ncol(o2CEscalation_no_mail_id_1),widths = 15)
writeData(wb,"o2CEscalation-1",o2CEscalation_no_mail_id_1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-2",cols = 1:ncol(o2CEscalation_no_mail_id_2),widths = 15)
writeData(wb,"o2CEscalation-2",o2CEscalation_no_mail_id_2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-3",cols = 1:ncol(o2CEscalation_no_mail_id_3),widths = 15)
writeData(wb,"o2CEscalation-3",o2CEscalation_no_mail_id_3,borders = "all",headerStyle = hs1)


path <- paste('./Output/',thisDate,"/Upload files/Unsent Emails on - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)





#No leads found in Crm---------------------------------------------------------------


wb <- createWorkbook()

addWorksheet(wb,"Reminder-1")
addWorksheet(wb,"Reminder-2")
addWorksheet(wb,"Reminder-3")
addWorksheet(wb,"Escalation-1")
addWorksheet(wb,"Escalation-2")
addWorksheet(wb,"Escalation-3")

addWorksheet(wb,"o2CReminder-1")
addWorksheet(wb,"o2CReminder-2")
addWorksheet(wb,"o2CReminder-3")
addWorksheet(wb,"o2CEscalation-1")
addWorksheet(wb,"o2CEscalation-2")
addWorksheet(wb,"o2CEscalation-3")

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

setColWidths(wb,"Escalation-1",cols = 1:ncol(LeadsNotFound_esc1),widths = 15)
writeData(wb,"Escalation-1",LeadsNotFound_esc1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-2",cols = 1:ncol(LeadsNotFound_esc2),widths = 15)
writeData(wb,"Escalation-2",LeadsNotFound_esc2,borders = "all",headerStyle = hs1)

setColWidths(wb,"Escalation-3",cols = 1:ncol(LeadsNotFound_esc3),widths = 15)
writeData(wb,"Escalation-3",LeadsNotFound_esc3,borders = "all",headerStyle = hs1)


setColWidths(wb,"o2CReminder-1",cols = 1:ncol(LeadsNotFound_o2CRem1),widths = 15)
writeData(wb,"o2CReminder-1",LeadsNotFound_o2CRem1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-2",cols = 1:ncol(LeadsNotFound_o2CRem2),widths = 15)
writeData(wb,"o2CReminder-2",LeadsNotFound_o2CRem2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CReminder-3",cols = 1:ncol(LeadsNotFound_o2CRem3),widths = 15)
writeData(wb,"o2CReminder-3",LeadsNotFound_o2CRem3,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-1",cols = 1:ncol(LeadsNotFound_o2CEsc1),widths = 15)
writeData(wb,"o2CEscalation-1",LeadsNotFound_o2CEsc1,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-2",cols = 1:ncol(LeadsNotFound_o2CEsc2),widths = 15)
writeData(wb,"o2CEscalation-2",LeadsNotFound_o2CEsc2,borders = "all",headerStyle = hs1)

setColWidths(wb,"o2CEscalation-3",cols = 1:ncol(LeadsNotFound_o2CEsc3),widths = 15)
writeData(wb,"o2CEscalation-3",LeadsNotFound_o2CEsc3,borders = "all",headerStyle = hs1)


path <- paste('./Output/',thisDate,"/Upload files/No Leads found in CRM - ",thisDate,'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)



#final_data ---------------------------------------------------------------------------
# 
# 
# 
# x<- data.frame(Emails=c("Cut and paste","Reminder-1","Reminder-2","Reminder-3","Escalation-1","Escalation-2","Escalation-3","02C Reminder Email","o2CReminder-1","o2CReminder-2","o2CReminder-3","o2CEscalation-1","o2CEscalation-2","o2CEscalation-3"), 
#                `Received Emails`=c(nrow(cut_and_paste),nrow(Reminder_1),nrow(Reminder_2),nrow(Reminder_3),nrow(Escalation_1),nrow(Escalation_2),nrow(Escalation_3),nrow(Reminder_02C)), 
#                `Sent Emails`=c(nrow(Cut_and_paste_dump1),nrow(Reminder_dump1),nrow(Reminder_dump2),nrow(Reminder_dump3),nrow(Escalation_dump1),nrow(Escalation_dump2),nrow(Escalation_dump3),nrow(Email_02c_dump)), 
#                `Not Sent Emails`=c(nrow(Cut_and_paste_no_mail_id),NROW(Reminder_no_mail_id_1),NROW(Reminder_no_mail_id_2),NROW(Reminder_no_mail_id_3),NROW(Escalation_no_mail_id_1),NROW(Escalation_no_mail_id_2),NROW(Escalation_no_mail_id_3),NROW(Email_02c_no_mail_id)),
#                `No leads found in DUMP`=c(0,nrow(LeadsNotFound_rem1),nrow(LeadsNotFound_rem2),nrow(LeadsNotFound_rem3),nrow(LeadsNotFound_esc1),nrow(LeadsNotFound_esc2),nrow(LeadsNotFound_esc3),0),
#                `02CReceived Emails`=c(nrow(o2CReminder_1),nrow(o2CReminder_2),nrow(o2CReminder_3),nrow(o2CEscalation_1),nrow(o2CEscalation_2),nrow(o2CEscalation_3),nrow(o2CReminder_02C)), 
#                `02CSent Emails`=c(nrow(o2CReminder_dump1),nrow(o2CReminder_dump2),nrow(o2CReminder_dump3),nrow(o2CEscalation_dump1),nrow(o2CEscalation_dump2),nrow(o2CEscalation_dump3),nrow(Email_02c_dump)), 
#                `02CNot Sent Emails`=c(NROW(o2CReminder_no_mail_id_1),NROW(o2CReminder_no_mail_id_2),NROW(o2CReminder_no_mail_id_3),NROW(o2CEscalation_no_mail_id_1),NROW(o2CEscalation_no_mail_id_2),NROW(o2CEscalation_no_mail_id_3),NROW(Email_02c_no_mail_id)),
#                `02CNo leads found in DUMP`=c(0,nrow(LeadsNotFound_o2CRem1),nrow(LeadsNotFound_o2CRem2),nrow(LeadsNotFound_o2CRem3),nrow(LeadsNotFound_o2CEsc1),nrow(LeadsNotFound_o2CEsc2),nrow(LeadsNotFound_o2CEsc3),0),
#                stringsAsFactors=FALSE)   
# 
# 
# final_data <- rbind(x, c("Total", colSums(x[,2:8])))
# 
# NewLeads <- final_data
# FilePath <- paste('./Output/',thisDate,"/Upload files/Overall Email sent Details - ",thisDate,'.xlsx',sep='')
# FileCreate <- FileCreation(NewLeads,FilePath)
# 

# upload_1 ----------------------------------------------------------------------


cut_up1 <- Cut_and_paste_dump1 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Nil",Comment="Cut and paste Email sent to customer",
                                                                    `Allocation Status`="Nil",`Alternate Acc No`="Nil")
Remi_up1 <-Reminder_dump1  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder one",Comment="Cut and Paste Reminder one Sent to customer",
                                                                   `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Remi_up2 <-Reminder_dump2  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder two",Comment="Cut and Paste Reminder two Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Remi_up3 <-Reminder_dump3  %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Cut and Paste Reminder three",Comment="Cut and Paste Reminder three Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Esc_up1 <-Escalation_dump1 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 1",Comment="Escalation Email 1 Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

Esc_up2 <-Escalation_dump2 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 2",Comment="Escalation Email 2 Sent to customer",

                                                                                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")
Esc_up3 <-Escalation_dump3 %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Escalation Email 3",Comment="Escalation Email 3 Sent to customer",
                                                                `Allocation Status`="Nil",`Alternate Acc No`="Nil")

# Rem_02C_up1 <-Email_02c_dump %>% select('CISserialno') %>% mutate(`Facility Status`="Nil",`CRM Status`="Nil",Comment="02C Reminder Email sent customer",
#                                                                   `Allocation Status`="Nil",`Alternate Acc No`="Nil")
# 
#Reminder_up1 <-rbind(cut_up1,Remi_up1,Remi_up2,Remi_up3,Rem_02C_up1)

Reminder_up1 <-rbind(cut_up1,Remi_up1,Remi_up2,Remi_up3)


if (nrow(Reminder_up1)>0) {
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Cut and paste & reminder Up1- ",nrow(Reminder_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}else{
  
  Reminder_up1 <- Reminder_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Cut and paste & reminder Up1-No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}

Esacaltion_up1 <-rbind(Esc_up1,Esc_up2,Esc_up3)


if (nrow(Esacaltion_up1)>0) {
  
  NewLeads <- Esacaltion_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation Up1- ",nrow(Esacaltion_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}else{
  
  Esacaltion_up1 <- Esacaltion_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- Esacaltion_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/Escalation Up1- No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


print("Completed")




o2CRemi_up1 <-o2CReminder_dump1  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="Cut and Paste o2CReminder one",Comment="Cut and Paste o2CReminder one Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CRemi_up2 <-o2CReminder_dump2  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="Cut and Paste o2CReminder two",Comment="Cut and Paste o2CReminder two Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CRemi_up3 <-o2CReminder_dump3  %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="Cut and Paste o2CReminder three",Comment="Cut and Paste o2CReminder three Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CEsc_up1 <-o2CEscalation_dump1 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="o2CEscalation Email 1",Comment="o2CEscalation Email 1 Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CEsc_up2 <-o2CEscalation_dump2 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="o2CEscalation Email 2",Comment="o2CEscalation Email 2 Sent to customer",
                                                                      
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")
o2CEsc_up3 <-o2CEscalation_dump3 %>% select('CISserialno') %>% mutate(`Facility Status`="02C  Customer sent Cut and Paste to Lender",`CRM Status`="o2CEscalation Email 3",Comment="o2CEscalation Email 3 Sent to customer",
                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

# o2CRem_02C_up1 <-Email_02c_dump %>% select('CISserialno') %>% mutate(`Facility Status`="Nil",`CRM Status`="Nil",Comment="02C o2CReminder Email sent customer",
#                                                                      `Allocation Status`="Nil",`Alternate Acc No`="Nil")

o2CReminder_up1 <-rbind(o2CRemi_up1,o2CRemi_up2,o2CRemi_up3)# o2CRem_02C_up1

if (nrow(o2CReminder_up1)>0) {
  
  NewLeads <- o2CReminder_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder Up1- ",nrow(o2CReminder_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}else{
  
  o2CReminder_up1 <- o2CReminder_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- o2CReminder_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CReminder Up1-No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}

o2CEsccaltion_up1 <-rbind(o2CEsc_up1,o2CEsc_up2,o2CEsc_up3)


if (nrow(o2CEsccaltion_up1)>0) {
  
  NewLeads <- o2CEsccaltion_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation Up1- ",nrow(o2CEsccaltion_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
}else{
  
  o2CEsccaltion_up1 <- o2CEsccaltion_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- o2CEsccaltion_up1
  FilePath <- paste('./Output/',thisDate,"/Upload files/o2CEscalation Up1- No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}


print("02C_Completed")


















