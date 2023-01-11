
rm(list = ls())
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_172')
library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(stringr)
library(mailR)
library (formattable)
library(scales)


options(scipen = 999)


setwd("C:\\R\\Current And Closed")

source('.\\Functionfile.R')


today_dat <- Sys.Date() 

date1 <-format(Sys.Date()-1,"%d%m")

thisDate = Sys.Date()
if(!dir.exists(paste(".\\Output\\",thisDate,sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,sep=""))
} 

if(!dir.exists(paste(".\\Output\\",thisDate,"\\Cut and paste",sep="")))
{
  dir.create(paste(".\\Output\\",thisDate,"\\Cut and paste",sep=""))
} 


#Read New cases dump --------------------------------


dump_path <- paste(".\\Input\\cis_dump_new_cases.csv",sep='')

dump <- fread(dump_path,na.strings = c('',NA))

dump <- dump[!is.na(dump$email_id),]

cut_and_paste <- dump %>% filter(LDB == "No" & prodstatus %in% c("CIS-WP-LD-220","CIS-WP-LP-310") & Facilitystatus == "00 To Start Action") %>% arrange(CISserialno)


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
FilePath <- path <- paste('./Output/',thisDate,"\\Cut and paste/Cut and paste - ",thisDate,'.xlsx',sep='')
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
Cut_and_paste_dump_fin <- dump1[!is.na(dump1$Email),] %>% distinct(CISserialno, .keep_all = T)

Cut_and_paste_dump1 <- Cut_and_paste_dump_fin %>% filter(Lender_Name !="Bajaj Finance" & Lender_Name !='SBI')
Cut_and_paste_dump2 <- Cut_and_paste_dump_fin %>% filter(Lender_Name =="Bajaj Finance")
Cut_and_paste_dump3 <- Cut_and_paste_dump_fin %>% filter(  Lender_Name =='SBI')
#mail function
if(nrow(Cut_and_paste_dump_fin)>0){
if(nrow(Cut_and_paste_dump1)>0){
lapply(Cut_and_paste_dump1$CISserialno, function(lead) {
  
#browser()
  accounts <- Cut_and_paste_dump1[Cut_and_paste_dump1$CISserialno == lead,]
  
  sender <-  'care@creditmantri.com'
  #recipients <-  accounts$Email_Id
  bcc_recipients <- 'Ciscare@creditmantri.com'
  message  <- str_interp('Reg : Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
  
  body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service (Credit Fit Program) for your account with ${accounts$Lender_Name}<br/><br/>You had subscribed for the above referenced service with CreditMantri. As part of this process we have already contacted your lender. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address in order to furnish the details on your account. We request you to cut and paste the below mentioned text and send it to the lender (${accounts$Email}) from your registered email address. Please copy us as cc in your email - care@creditmantri.com <br/> <br/> <b>Kindly ignore this mail if you have already sent the below text to the lender.</b><br/><br/>Please forward us any responses that you have received or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/>')
  
  body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Subject: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>   On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a  ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible. <br/><br/>Thanks and Regards <br/>${accounts$Customer_Name}')
  
  body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
  
  email_body <- paste(body1, body2, body3)
  
  recipients <- accounts$Email_Id
  
  mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body, filename = NULL)
  
  print(paste("Cut and paste -",accounts$Account_No))
  
})

}

if(nrow(Cut_and_paste_dump2)>0){
#mail function

lapply(Cut_and_paste_dump2$CISserialno, function(lead){
  
  #browser()
  accounts <- Cut_and_paste_dump2[Cut_and_paste_dump2$CISserialno == lead,]
  
  sender <-  'care@creditmantri.com'
  #recipients <-  accounts$Email_Id
  bcc_recipients <- 'Ciscare@creditmantri.com'
  message  <- str_interp('Reg : Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
  
  body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service (Credit Fit Program) for your account with ${accounts$Lender_Name}<br/><br/>You had subscribed for the above referenced service with CreditMantri. As part of this process we have already contacted your lender. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address in order to furnish the details on your account.<br/> 1.Send Bajaj Finance an email - details mentioned below
<br/> 2.Send a PDF copy of your account statement to us.(care@creditmantri.com). If you do not have it, please request the same from the lender. You may do so in the same email as well. <br/>For the email to your lender, please cut and paste the below mentioned text and send it to the lender (wecare@bajajfinserv.in) from your registered email address.
Kindly retain the subject line. Please copy us as cc in your email - care@creditmantri.com<br/><b>Kindly ignore this mail if you have already sent the below text to the lender.</b><br/>Please forward us any responses that you have received or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/>')
  
  body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Subject: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>   On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a  ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.<br/><br/>(You may add this line if you also need the e-statement from Bajaj Finance) <br/>Request you to also send me a PDF copy of my latest account statement to this email address.')
  
  body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
  
  email_body <- paste(body1, body2, body3)
  
  recipients <- accounts$Email_Id
  
  mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body, filename = NULL)
  
  print(paste("Cut and paste -",accounts$Account_No))
  
})
}

if(nrow(Cut_and_paste_dump3)>0){
  lapply(Cut_and_paste_dump3$CISserialno, function(lead){
    
    #browser()
    accounts <- Cut_and_paste_dump3[Cut_and_paste_dump3$CISserialno == lead,]
    
    sender <-  'care@creditmantri.com'
    #recipients <-  accounts$Email_Id
    bcc_recipients <- 'Ciscare@creditmantri.com'
    message  <- str_interp('Reg : Need Outstanding Amount for ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}')
    
    body1 <- str_interp('<br/>Dear ${accounts$Customer_Name},<br/><br/>Ref : Subscription to our Credit Improvement Service for your account with ${accounts$Lender_Name}<br/><br/>You had subscribed for the above referenced service with CreditMantri. As part of this process we have already contacted your lender. As part of their information security policy and for audit purpose, they have requested an email from your registered E-Mail address in order to furnish the details on your account. We request you to:<br/>1. Send SBI an email - details mentioned below
<br/>2. Provide your bank branch name, location and contact number to us.(care@creditmantri.com). If you do not have it, please request the same from the lender. You may do so in the same email as referenced in point 1.
<br/>For the email to your lender, please cut and paste the below mentioned text and send it to the lender (contactcentre@sbi.co.in) from your registered email address.
Kindly retain the subject line. Please copy us as cc in your email - care@creditmantri.com
<br/><b>Kindly ignore this mail if you have already sent the below text to the lender.</b><br/>
Please forward us any responses that you have received or you may receive from the bank to care@creditmantri.com to further proceed with your account resolution.<br/>')
    
    body2 <- str_interp('<br/>---------------------------------------Cut Here---------------------------------------<br/> <br/>  Dear Sir/Madam, <br/> <br/>Subject: ${accounts$Lender_Name} ${accounts$Account_type} ${accounts$Account_No}<br/><br/>   On a scrutiny of my Equifax bureau report, I observe that the referenced account has been reported with a  ${accounts$Account_status} status in the repayment history. This is hindering my ability to avail fresh loans and I would like to know what needs to be done to update the account status as CLOSED/STANDARD. I need your assistance to resolve this issue with the bureau and help me become loan eligible.<br/><br/>(You may add this line if you also need the Bank contact details from SBI) <br/>
Request you to also send me the contact details of my branch, so that I can get in touch with them re: the resolution of my issue.
<br/>Branch Name: (please provide your branch name)
<br/>Location. (please provide your location name)
 <br/>. <br/><br/>Thanks and Regards <br/>${accounts$Customer_Name}')
    body3 <- '<br/><br/>---------------------------------------Cut Here---------------------------------------<br/> <br/> Kindly forward any response/feedback that you may receive from them.<br/><br />Thanks and Regards <br />Team CreditMantri'
    
    email_body <- paste(body1, body2, body3)
    
    recipients <- accounts$Email_Id
    
    mail_function(sender, recipients, cc_recipients = NULL,bcc_recipients,message, email_body, filename = NULL)
    
    print(paste("Cut and paste -",accounts$Account_No))
    
  })
}
}
}else{
  cut_and_paste <- cut_and_paste %>%select('CISserialno','Date','Customer_Name','Lender_Name','Account_type','Account_No',
                                     'Account_status','Facilitystatus','Email_Id','Phone_Number','Address','PAN','dob','Product_Status',
                                     'LDB','Flag')
  
  NewLeads <- cut_and_paste
  FilePath <- paste('./Output/',thisDate,"\\Cut and paste/Cut and paste - No Leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
  
  }


#unsent Email ----------------------------------------------------------------------------------------

wb <- createWorkbook()
addWorksheet(wb,"Cut and paste")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")

setColWidths(wb,"Cut and paste",cols = 1:ncol(Cut_and_paste_no_mail_id),widths = 15)
writeData(wb,"Cut and paste",Cut_and_paste_no_mail_id,borders = "all",headerStyle = hs1)

path <- paste('./Output/',thisDate,"\\Cut and paste/cut and paste Unsent Emails on - ",thisDate,'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)





#upload_1 ----------------------------------------------------------------------

cut_up1 <- Cut_and_paste_dump_fin %>% select('CISserialno') %>% mutate(`Facility Status`="01C Requested customer to send Mail",`CRM Status`="Nil",Comment="Cut and paste Email sent to customer",
                                                                    `Allocation Status`="Nil",`Alternate Acc No`="Nil")
Reminder_up1 <-cut_up1

if (nrow(Reminder_up1)>0) {
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"\\Cut and paste/Cut and paste Up1- ",nrow(Reminder_up1)," - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}else{
  
  Reminder_up1 <- Reminder_up1 %>% mutate(CISserialno="",`Facility Status`="",`CRM Status`="",Comment="",`Allocation Status`="",`Alternate Acc No`="")
  
  NewLeads <- Reminder_up1
  FilePath <- paste('./Output/',thisDate,"\\Cut and paste/Cut and paste Up1-No leads found - ",thisDate,'.xlsx',sep='')
  FileCreate <- FileCreation(NewLeads,FilePath)
  
}

print("Completed")





















