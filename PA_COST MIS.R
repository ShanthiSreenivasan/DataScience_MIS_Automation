rm(list=ls())
library(magrittr)
library(tibble)
library(dplyr)
library(plyr)
library(tidyr)
library(stringr)
library(lubridate)
library(openxlsx)
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

setwd("C:\\R\\PA_Cost")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 



##################################################
#View(CIS)
engagement_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Daily_Engagment_Stats_DUMP.csv")



Login_dump <- read.xlsx("C:\\R\\PA_Cost\\Input\\Login_details.xlsx",sheet='Login_detail')

del_dump <- read.xlsx("C:\\R\\PA_Cost\\Input\\Login_details.xlsx", sheet='Auto_R')

merge_dump<-left_join(engagement_dump,del_dump,by= c('NODEID','name'))

df1_chk<-merge_dump %>% filter(!DEL%in% c(1))


daily_enga_dump<-df1_chk
names(merge_dump)
# daily_enga_dump <-  daily_enga_dump %>% 
#   mutate(
#     Particulars = case_when(
#       grepl('IDPP-040', name, ignore.case = T) ~ 'IDPP-40',
#       grepl('IDPP-050', name, ignore.case = T) ~ 'IDPP-50',
#       grepl('-050DROPOFF|0045DROPOFF|Status-50|Status-45', name, ignore.case = T) ~ '45-50',
#       grepl('CIS', name, ignore.case = T) ~ 'CIS',
#       grepl('BHR-RED-DROP|BHRL-RED-DROP',name, ignore.case = T) ~ 'BHR-RED-Dropoff',
#       grepl('CHR-RED-DROP|CHRL-RED-DROP',name, ignore.case = T) ~ 'CHR-RED-Dropoff',
#       grepl('CIS-RED-DROP',name, ignore.case = T) ~ 'CIS-RED-Dropoff',
#       grepl('BHR-GREEN-DROP|BHRL-GREEN-DROP|BYS-DROP',name, ignore.case = T) ~ 'BHR-GREEN-Dropoff',
#       grepl('CHR-GREEN-DROP|CHRL-GREEN-DROP',name, ignore.case = T) ~ 'CHR-GREEN-Dropoff',
#       grepl('CIS-GREEN-DROP',name, ignore.case = T) ~ 'CIS-GREEN-Dropoff',
#       grepl('BHR-GREEN|BHR-RENEWAL-GREEN|BHR-TEST|BHRL-GREEN|BHRL-RENEWAL-GREEN|BHRL-TEST|BHR-RENEWAL',name, ignore.case = T) ~ 'BHR-GREEN',
#       grepl('BHR-RED|BHRL-RED',name, ignore.case = T) ~ 'BHR-RED',
#       grepl('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST',name, ignore.case = T) ~ 'CHR-GREEN',
#       grepl('CHR-RED',name, ignore.case = T) ~ 'CHR-RED',
#       grepl('Acko|-ISC-',name, ignore.case = T) ~ 'Acko',
#       name %in% c('PAYU') ~ 'PayU',
#       grepl('APPDOWNLOAD', name, ignore.case = T) ~ 'APPDOWNLOAD',
#       grepl('RFC|DROP', name, ignore.case = T) ~ 'Referrals',
#       grepl('CFP', name, ignore.case = T) ~ 'CFP',
#       grepl('SBC_PBC|SBC|PBC', name, ignore.case = T) ~ 'PBC-SBC',
#       grepl('LBC', name, ignore.case = T) ~ 'LenderBase',
#       name %in% c('RED-REPRO') ~ 'RED-Profiling',
#       grepl('CHR', name, ignore.case = T) ~ 'CHR-GREEN',
#       grepl('TEST', name, ignore.case = T) ~ 'PBC-SBC'
#     )
#   )
# 

creden<-daily_enga_dump

creden$Particulars<-Login_dump$Tag[match(creden$flow_id, Login_dump$flow_id)]



#View(creden)

# Replace String with Another Stirng
creden[creden == 'email'] <- 'Auto._Email'

creden[creden == 'smstrans'] <- 'Auto._SMS'

creden[creden == 'whatsapp'] <- 'Auto._Watsapp'

creden[creden == 'voice'] <- 'IVR'


#str_replace_all(creden$CHANNEL, "smstrans", "Auto._SMS")

#str_replace_all(creden$CHANNEL, "whatsapp", "Auto._Watsapp")


creden<-creden %>% filter(!Particulars=='NA')
#names(creden)


#names(creden)
#View(creden)
wo_reload<-creden %>% filter(!Particulars=='Reupload',CHANNEL == "Auto._Watsapp") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_Watsapp_Sent = sum(Attempted), ACL_Watsapp_Cost=ACL_Watsapp_Sent*0.51)

cred_wats<-wo_reload


wo_reload_DF<-creden %>% filter(!Particulars=='Reupload',CHANNEL == "Auto._Watsapp") %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Attempted), Cost=Sent*0.51)

cred_wats_DF<-wo_reload_DF


Credence_Watsapp<-creden %>% filter(!Particulars=='Reupload',CHANNEL == "Auto._Watsapp")%>% dplyr::summarise(ACL_Watsapp_Sent = sum(Attempted), ACL_Watsapp_Cost=ACL_Watsapp_Sent*0.51)


cred_email<- creden %>% filter(CHANNEL == "Auto._Email") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_Email_Sent = sum(Attempted),ACL_Email_Cost=ACL_Email_Sent*0.01)

cred_email_DF<- creden %>% filter(CHANNEL == "Auto._Email") %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Attempted),Cost=Sent*0.01)


Credence_Email<-creden %>% filter(CHANNEL == "Auto._Email") %>% dplyr::summarise(ACL_Email_Sent = sum(Attempted),ACL_Email_Cost=ACL_Email_Sent*0.01)

#cred_sms<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1,2)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms01<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms2<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(2)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = 2*sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms<-rbind(cred_sms01,cred_sms2)


cred_sms01_DF<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1)) %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Attempted),Cost=Sent*0.12)

cred_sms2_DF<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(2)) %>% group_by(Particulars) %>% dplyr::summarise(Sent = 2*sum(Attempted),Cost=Sent*0.12)

cred_sms_DF<-rbind(cred_sms01_DF,cred_sms2_DF)






#Credence_SMS<- creden %>% filter(CHANNEL == "Auto._SMS") %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

Credence_SMS01<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1)) %>% dplyr::summarise(ACL_SMS1 = sum(Attempted),Cost1=ACL_SMS1*0.12)
Credence_SMS2<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(2)) %>% dplyr::summarise(ACL_SMS2 = 2*sum(Attempted),Cost2=ACL_SMS2*0.12)


ACL_SMS_Sent<-Credence_SMS01$ACL_SMS1+Credence_SMS2$ACL_SMS2

ACL_SMS_Cost<-Credence_SMS01$Cost1+Credence_SMS2$Cost2

Credence_SMS<-cbind(ACL_SMS_Sent,ACL_SMS_Cost)

cred_ivr<- creden %>% filter(CHANNEL == "IVR") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_IVR_Sent = sum(Delivered),ACL_IVR_Cost=ACL_IVR_Sent*0.22)

cred_ivr_DF<- creden %>% filter(CHANNEL == "IVR") %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Delivered),`Cost`=Sent*0.22)

Credence_IVR<- creden %>% filter(CHANNEL == "IVR") %>% dplyr::summarise(ACL_IVR_Sent = sum(Delivered),ACL_IVR_Cost=ACL_IVR_Sent*0.22)


#Credence_ivr<-rbind(cred_v1,cred_v2)

#Cred_Panel<-comma(Cred_Panel)
################################################SmartPing_Individual_IVR

#SPivr_dump <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP.csv") %>% mutate(`CHANNEL`="IVR")




setwd("C:\\R\\PA_Cost\\Input\\SmartPing_IVR")
getwd()
library(data.table)
SPivr_dump <- 
  list.files(path = "C:\\R\\PA_Cost\\Input\\SmartPing_IVR", pattern = "*.csv") %>% 
  map_df(~read_csv(.))

write.csv(SPivr_dump,file = paste0("C:\\R\\PA_Cost\\SPivr_dump", Sys.Date()-1, '.csv'))

names(SPivr_dump)

SPivr_dump<-SPivr_dump %>% mutate(`CHANNEL`="IVR")

#SPivr_dump1 <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP1.csv") %>% mutate(`CHANNEL`="IVR")

#SPivr_dump2 <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP2.csv") %>% mutate(`CHANNEL`="IVR")

#SPivr_df1<-rbind(SPivr_dump1,SPivr_dump2)


#SPivr_df2<-rbind(SPivr_df1,SPivr_dump3)

#SPivr_dump<-SPivr_df2

SPivr_dump<-SPivr_dump %>% filter(grepl('CALL_SUCCESS', Overcall_Call_status, ignore.case = T))

SPivr_dump <-SPivr_dump %>%
  mutate(Duration_flag = case_when(
    Duration <= 15 ~ 1,
    Duration > 15 & Duration <= 30 ~ 2,
    Duration > 30 & Duration <= 45 ~ 3,
    Duration > 45 & Duration <= 60 ~ 4,
    Duration > 60 & Duration <= 75 ~ 5
  ))

SPivr_dump <-  SPivr_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('SB_PB', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase'
      
    )
  )

SPivr_dump$Particulars[is.na(SPivr_dump$Particulars)] <- 'PBC-SBC'

SPivr <- SPivr_dump

SPivr<-SPivr %>% filter(grepl('CALL_SUCCESS', Overcall_Call_status, ignore.case = T)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_IVR_Sent` = sum(Duration_flag),Smartping_IVR_Cost=Smartping_IVR_Sent*0.12)


SPivr_DF<-SPivr_dump %>% filter(grepl('CALL_SUCCESS', Overcall_Call_status, ignore.case = T)) %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Duration_flag),Cost=Sent*0.12)

Smartping_IVR<- SPivr_dump %>% dplyr::summarise(`Smartping_IVR_Sent` = sum(Duration_flag),Smartping_IVR_Cost=Smartping_IVR_Sent*0.12)


##########################SMSFolio

SMSFolio_dump <- read.csv("C:\\R\\PA_Cost\\Input\\SMSFolio_CreditMantri_Summary_MTD.csv") %>% mutate(`CHANNEL`="SMS")

names(SMSFolio_dump)

SMSFolio_dump <- SMSFolio_dump %>% distinct(Account, CampaignName,TotalSent, .keep_all= TRUE)


SMSFolio_dump <-  SMSFolio_dump %>% 
  mutate(
    Particulars = case_when(
      CampaignName %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      CampaignName %in% c('REPROFILING-GREEN|GREEN-REPRO') ~ 'GREEN-Profiling',
      CampaignName %in% c('CIS-REP') ~ 'CIS-Reprofiling',
      grepl('REPROFILING-RED|TEST-REPRO|RED-REPRO', CampaignName, ignore.case = T) ~ 'RED-Profiling',
      grepl('REPROFILING-GREEN|GREEN-REPRO', CampaignName, ignore.case = T) ~ 'GREEN-Profiling',
      Account %in% c('creditmantri2') ~ 'Acko',
      Account %in% c('creditmantri3') & grepl('LBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('SBC_PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      Account %in% c('creditmantri4') ~ 'LenderBase'
    )
  )

SMSFolio_dump$Particulars[is.na(SMSFolio_dump$Particulars)] <- 'PBC-SBC'



SMSFolio1 <-  SMSFolio_dump %>% filter(Account %in% c('creditmantri2','creditmantri3')) %>% 
  mutate(`Total`=TotalSent)

SMSFolio2 <-  SMSFolio_dump %>% filter(Account %in% c('creditmantri4')) %>% 
  mutate(`Total`=TotalSent-(UniqueSent-(UniqueDelivered+Undelivered)))


SMSFolio<-bind_rows(SMSFolio1,SMSFolio2)# %>% reduce(full_join,by=c("Particulars")) %>% adorn_totals("row")# %>% dplyr::rename(`Spent` = 'Total')

smsfolio<-SMSFolio %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`SMSFolio_SMS_Sent` = sum(Total),`SMSFolio_SMS_Cost`=SMSFolio_SMS_Sent*0.115)

smsfolio_DF<-SMSFolio %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Total),Cost=Sent*0.115)


SMSFolio<- SMSFolio %>% dplyr::summarise(`SMSFolio_SMS_Sent` = sum(Total),SMSFolio_SMS_Cost=SMSFolio_SMS_Sent*0.115)




################################################SmartPing_SMS

SPsms_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Smartping_CreditMantri_Summary_MTD.csv") %>% mutate(`CHANNEL`="SMS")

#SPsms_dump <- read.xlsx("C:\\R\\PA_Cost\\Input\\Smartping_SMS_DUMP.xlsx", sheet='CampaignData') %>% mutate(`CHANNEL`="SMS")

#SPsms_login_dump <- read.xlsx("C:\\R\\PA_Cost\\Input\\SMARTPING_SMS_login_dump.xlsx")

#names(SPsms_dump)
#SPsms_dump <- SPsms_dump %>% distinct(Account, CampaignName,TotalSent, .keep_all= TRUE)

SPsms_dump$TotalSent <- as.numeric(SPsms_dump$TotalSent)

SPsms_dump$UniqueDelivered <- as.numeric(SPsms_dump$UniqueDelivered)

SPsms_dump$Undelivered <- as.numeric(SPsms_dump$Undelivered)

SPsms_dump <-  SPsms_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('RFC-50-DROPOFF|RFC-50-', CampaignName, ignore.case = T) ~ 'Status 45-50',
      grepl('Acko|-ISC--',CampaignName, ignore.case = T) ~ 'Acko',
      grepl('www.Acko.com', URL, ignore.case = T) ~ 'Acko',
      grepl('PAYU|payu',CampaignName, ignore.case = T) ~ 'PayU',
      CampaignName %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      CampaignName %in% c('REPROFILING-GREEN|GREEN-REPRO') ~ 'GREEN-Profiling',
      CampaignName %in% c('CIS-REP') ~ 'CIS-Reprofiling',
      grepl('REPROFILING-RED|TEST-REPRO|RED-REPRO', CampaignName, ignore.case = T) ~ 'RED-Profiling',
      grepl('REPROFILING-GREEN|GREEN-REPRO', CampaignName, ignore.case = T) ~ 'GREEN-Profiling',
      grepl('CIS-REP', CampaignName, ignore.case = T) ~ 'CIS-Reprofiling',
      grepl('Auto_APPDOWNLOAD', CampaignName, ignore.case = T) ~ 'Auto_APPDOWNLOAD',
      grepl('APPDOWNLOAD|WEBLINK|APPDWN', CampaignName, ignore.case = T) ~ 'APPDOWNLOAD',
      grepl('BHR-RED-DROP|BHRL-RED-DROP',CampaignName, ignore.case = T) ~ 'BHR-RED-Dropoff',
      grepl('CHR-RED-DROP|CHRL-RED-DROP',CampaignName, ignore.case = T) ~ 'CHR-RED-Dropoff',
      grepl('CIS-RED-DROP',CampaignName, ignore.case = T) ~ 'CIS-RED-Dropoff',
      grepl('BHR-GREEN-DROP|BHRL-GREEN-DROP|BYS-DROP',CampaignName, ignore.case = T) ~ 'BHR-GREEN-Dropoff',
      grepl('CHR-GREEN-DROP|CHRL-GREEN-DROP',CampaignName, ignore.case = T) ~ 'CHR-GREEN-Dropoff',
      grepl('CIS-GREEN-DROP',CampaignName, ignore.case = T) ~ 'CIS-GREEN-Dropoff',
      grepl('CHR-RED|CHR-RENEWAL-RED', CampaignName, ignore.case = T) ~ 'CHR-RED',
      grepl('CIS|cis-test-', CampaignName, ignore.case = T) ~ 'CIS',
      CampaignName %in% c('CIS') ~ 'CIS',
      grepl('RFC|rfc-test-', CampaignName, ignore.case = T) ~ 'Referrals',
      CampaignName %in% c('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST') ~ 'CHR-GREEN',
      grepl('CFP|cfp-test-|TEST-CFP-', CampaignName, ignore.case = T) ~ 'CFP',
      CampaignName %in% c('SBC_PBC','SBC','PBC') ~ 'PBC-SBC',
      #grepl('SBC_PBC|SBC|PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      grepl('CHR|chr-test', CampaignName, ignore.case = T) ~ 'CHR-GREEN',
      
    )
  )

SPsms_dump$Particulars[is.na(SPsms_dump$Particulars)] <- 'PBC-SBC'

#SPsms_dump$Particulars<-SPsms_login_dump$Particulars[match(SPsms_dump$CampaignName, SPsms_login_dump$CampaignName)]

#SPsms_dump$Particulars[is.na(SPsms_dump$Particulars)] <- 'PBC-SBC'

SPsms_dump <-  SPsms_dump %>% mutate(`Total`= TotalSent - (UniqueDelivered+Undelivered))


SPsmsU<-SPsms_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_SMS_UniqueDeli._Sent` = sum(UniqueDelivered),`Smartping_SMS_UniqueDeli._Cost`=Smartping_SMS_UniqueDeli._Sent*0.105)


SPsmsU_DF<-SPsms_dump %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(UniqueDelivered),Cost=Sent*0.105)


SPSMSU<- SPsms_dump %>% dplyr::summarise(`Smartping_SMS_UniqueDeli._Sent` = sum(UniqueDelivered),`Smartping_SMS_UniqueDeli._Cost`=Smartping_SMS_UniqueDeli._Sent*0.105)

#View(SPsms_dump)
SPsms<-SPsms_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_SMS_Sent` = sum(Total+Undelivered),`Smartping_SMS_Cost`=Smartping_SMS_Sent*0.02)

SPsms_DF<-SPsms_dump %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Total+Undelivered),Cost=Sent*0.02)


SPSMS<- SPsms_dump %>% dplyr::summarise(`Smartping_SMS_Sent` = sum(Total+Undelivered),`Smartping_SMS_Cost`=Smartping_SMS_Sent*0.02)

SP_SMS_Sent<-sum(SPSMSU$Smartping_SMS_UniqueDeli._Sent+SPSMS$Smartping_SMS_Sent)

SP_SMS_Cost<-sum(SPSMSU$Smartping_SMS_UniqueDeli._Cost+SPSMS$Smartping_SMS_Cost)

SP_SMS<-cbind(SP_SMS_Sent,SP_SMS_Cost) %>% as.data.frame()


write.xlsx(SPsms_dump, file = paste0("C:\\R\\PA_Cost\\Output\\SP_SMS-", Sys.Date()-1, '.xlsx'))

write.xlsx(SMSFolio_dump, file = paste0("C:\\R\\PA_Cost\\Output\\SMS_FOLIO_SMS-", Sys.Date()-1, '.xlsx'))


################################################Mailkoot


kasp_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Kasplo_CreditMantri_Summary_MTD.csv") %>% mutate(`CHANNEL`="email")



kasp_dump <-  kasp_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('RFC-50-DROPOFF|RFC-50-', CampaignName, ignore.case = T) ~ 'Status 45-50',
      grepl('Acko|-ISC-',CampaignName, ignore.case = T) ~ 'Acko',
      grepl('www.Acko.com', URL, ignore.case = T) ~ 'Acko',
      CampaignName %in% c('PAYU') ~ 'PayU',
      CampaignName %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      CampaignName %in% c('REPROFILING-GREEN|GREEN-REPRO') ~ 'GREEN-Profiling',
      CampaignName %in% c('CIS-REP') ~ 'CIS-Reprofiling',
      grepl('REPROFILING-RED|TEST-REPRO|RED-REPRO', CampaignName, ignore.case = T) ~ 'RED-Profiling',
      grepl('REPROFILING-GREEN|GREEN-REPRO', CampaignName, ignore.case = T) ~ 'GREEN-Profiling',
      grepl('CIS-REP', CampaignName, ignore.case = T) ~ 'CIS-Reprofiling',
      
      grepl('Auto_APPDOWNLOAD', CampaignName, ignore.case = T) ~ 'Auto_APPDOWNLOAD',
      grepl('APPDOWNLOAD|APPDWN', CampaignName, ignore.case = T) ~ 'APPDOWNLOAD',
      grepl('BHR-RED-DROP|BHRL-RED-DROP',CampaignName, ignore.case = T) ~ 'BHR-RED-Dropoff',
      grepl('CHR-RED-DROP|CHRL-RED-DROP',CampaignName, ignore.case = T) ~ 'CHR-RED-Dropoff',
      grepl('CIS-RED-DROP',CampaignName, ignore.case = T) ~ 'CIS-RED-Dropoff',
      grepl('BHR-GREEN-DROP|BHRL-GREEN-DROP|BYS-DROP',CampaignName, ignore.case = T) ~ 'BHR-GREEN-Dropoff',
      grepl('CHR-GREEN-DROP|CHRL-GREEN-DROP',CampaignName, ignore.case = T) ~ 'CHR-GREEN-Dropoff',
      grepl('CIS-GREEN-DROP',CampaignName, ignore.case = T) ~ 'CIS-GREEN-Dropoff',
      grepl('CHR-RED|CHR-RENEWAL-RED',CampaignName, ignore.case = T) ~ 'CHR-RED',
      grepl('BHR-RED|BHRL-RED|BHR-RENEWAL-RED',CampaignName, ignore.case = T) ~ 'BHR-RED',
      grepl('RFC', CampaignName, ignore.case = T) ~ 'Referrals',
      grepl('CIS', CampaignName, ignore.case = T) ~ 'CIS',
      grepl('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST',CampaignName, ignore.case = T) ~ 'CHR-GREEN',
      grepl('BHR-GREEN|BHR-RENEWAL-GREEN|BHR-TEST|BHRL-GREEN|BHRL-RENEWAL-GREEN|BHRL-TEST',CampaignName, ignore.case = T) ~ 'BHR-GREEN',
      grepl('CFP', CampaignName, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      grepl('CHR|CHR-GREEN', CampaignName, ignore.case = T) ~ 'CHR-GREEN',
      grepl('TEST', CampaignName, ignore.case = T) ~ 'PBC-SBC'
      
    )
  )
kasp_dump$Particulars[is.na(kasp_dump$Particulars)] <- 'Referrals'

mailkoot<-kasp_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Mailkoot_Email_Sent` = sum(UniqueSent),`Mailkoot_Email_Cost`=Mailkoot_Email_Sent*0.009)


mailkoot_DF<-kasp_dump %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(UniqueSent),Cost=Sent*0.009)


MailKoot<- kasp_dump %>% dplyr::summarise(`Mailkoot_Email_Sent` = sum(UniqueSent),Mailkoot_Email_Cost=Mailkoot_Email_Sent*0.009)







################################################Netcore smart


net_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Email_summary_report.csv") %>% mutate(`CHANNEL`="email")

net_dump <-  net_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('RFC-50-DROPOFF|RFC-50-', Forward, ignore.case = T) ~ 'Status 45-50',
      grepl('Acko|-ISC-',Forward, ignore.case = T) ~ 'Acko',
      Forward %in% c('PAYU') ~ 'PayU',
      Forward %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      Forward %in% c('REPROFILING-GREEN|GREEN-REPRO') ~ 'GREEN-Profiling',
      Forward %in% c('CIS-REP') ~ 'CIS-Reprofiling',
      grepl('REPROFILING-RED|TEST-REPRO|RED-REPRO', Forward, ignore.case = T) ~ 'RED-Profiling',
      grepl('REPROFILING-GREEN|GREEN-REPRO', Forward, ignore.case = T) ~ 'GREEN-Profiling',
      grepl('CIS-REP', Forward, ignore.case = T) ~ 'CIS-Reprofiling',
      
      grepl('Auto_APPDOWNLOAD', Forward, ignore.case = T) ~ 'Auto_APPDOWNLOAD',
      grepl('APPDOWNLOAD|WEBLINK|APPDWN', Forward, ignore.case = T) ~ 'APPDOWNLOAD',
      grepl('BHR-RED-DROP|BHRL-RED-DROP',Forward, ignore.case = T) ~ 'BHR-RED-Dropoff',
      grepl('CHR-RED-DROP|CHRL-RED-DROP',Forward, ignore.case = T) ~ 'CHR-RED-Dropoff',
      grepl('CIS-RED-DROP',Forward, ignore.case = T) ~ 'CIS-RED-Dropoff',
      grepl('BHR-GREEN-DROP|BHRL-GREEN-DROP|BYS-DROP',Forward, ignore.case = T) ~ 'BHR-GREEN-Dropoff',
      grepl('CHR-GREEN-DROP|CHRL-GREEN-DROP',Forward, ignore.case = T) ~ 'CHR-GREEN-Dropoff',
      grepl('CIS-GREEN-DROP',Forward, ignore.case = T) ~ 'CIS-GREEN-Dropoff',
      grepl('CHR-RED|CHR-RENEWAL-RED',Forward, ignore.case = T) ~ 'CHR-RED',
      grepl('BHR-RED|BHRL-RED|BHR-RENEWAL-RED',Forward, ignore.case = T) ~ 'BHR-RED',
      grepl('CIS', Forward, ignore.case = T) ~ 'CIS',
      grepl('RFC', Forward, ignore.case = T) ~ 'Referrals',
      grepl('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST',Forward, ignore.case = T) ~ 'CHR-GREEN',
      grepl('BHR-GREEN|BHR-RENEWAL-GREEN|BHR-TEST|BHRL-GREEN|BHRL-RENEWAL-GREEN|BHRL-TEST',Forward, ignore.case = T) ~ 'BHR-GREEN',
      grepl('CFP', Forward, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC', Forward, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', Forward, ignore.case = T) ~ 'LenderBase',
      grepl('CHR|CHR-GREEN', Forward, ignore.case = T) ~ 'CHR-GREEN',
      grepl('TEST', Forward, ignore.case = T) ~ 'PBC-SBC'
      
    )
  )

net_dump$Particulars[is.na(net_dump$Particulars)] <- 'PBC-SBC'

netcore<-net_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Netcore_smart_Email_Sent` = sum(Total.Published),`Netcore_smart_Email_Cost`=Netcore_smart_Email_Sent*0.0125)

netcore_DF<-net_dump %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Total.Published),Cost=Sent*0.0125)

NETCORE<- net_dump %>% dplyr::summarise(`Netcore_smart_Email_Sent` = sum(Total.Published),Netcore_smart_Email_Cost=Netcore_smart_Email_Sent*0.0125)

#############################Karix

#Activity_summary_1 <- read.csv("C:\\R\\PA_Cost\\Input\\Activity_summary_1.csv",sep=",") %>% mutate(`CHANNEL`="SMS")

#Activity_summary_1 <- read.xlsx("C:\\R\\PA_Cost\\Input\\Activity_summary_1.xlsx") %>% select(File.Name,Total.Records) %>% mutate(`CHANNEL`="SMS")


#Activity_summary_2 <- read.csv("C:\\R\\PA_Cost\\Input\\Activity_summary_2.csv",sep="") %>% mutate(`CHANNEL`="SMS")

#Activity_summary_2 <- read.xlsx("C:\\R\\PA_Cost\\Input\\Activity_summary_2.xlsx") %>% select(File.Name,Total.Records) %>% mutate(`CHANNEL`="SMS")


#karix_dump<-rbind(Activity_summary_1,Activity_summary_2)
#SHOULD BE IN XLSX FILE FORMAT

#Activity_summary <- read.xlsx("C:\\R\\PA_Cost\\Input\\Activity_summary.xlsx") %>% select(File.Name,Total.Records) %>% mutate(`CHANNEL`="SMS")
Activity_summary <- read.xlsx("C:\\R\\PA_Cost\\Input\\CampaignsReport_CampaignWiseSummaryData.xlsx") %>% select(FileName,Total.Records) %>% mutate(`CHANNEL`="SMS")

karix_dump <-Activity_summary
karix_dump<-karix_dump %>% filter(!FileName=='') 
karix_dump$Total.Records<-as.numeric(karix_dump$Total.Records)

karix_dump <-  karix_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('LB|LBC', FileName, ignore.case = T) ~ 'LenderBase',
      (!is.na(FileName) & !grepl('LB|LBC', FileName, ignore.case = T)) ~ 'PBC-SBC'
      
    )
  )

karix_dump$Particulars[is.na(karix_dump$Particulars)] <- 'PBC-SBC'

karix<-karix_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Karix_SMS_Sent` = sum(Total.Records),`Karix_SMS_Cost`=Karix_SMS_Sent*0.125)

karix_DF<-karix_dump %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(Total.Records),Cost=Sent*0.125)

KARIX<- karix_dump %>% dplyr::summarise(`Karix_SMS_Sent` = sum(Total.Records),`Karix_SMS_Cost`=Karix_SMS_Sent*0.125)

###########################

SummaryReport <- read.csv("C:\\R\\PA_Cost\\Input\\SummaryReport.csv") %>% mutate(`CHANNEL`="IVR")#%>% select(Campaign.Name,Total.Calls) 


karix_dump_IVR <-SummaryReport
karix_dump_IVR<-karix_dump_IVR %>% filter(!Campaign.Name=='') 
karix_dump_IVR$X15Sec.Pulse<-as.numeric(karix_dump_IVR$X15Sec.Pulse)
names(karix_dump_IVR)

karix_dump_IVR <-  karix_dump_IVR %>% 
  mutate(
    Particulars = case_when(
      grepl('LBC|LB', Campaign.Name, ignore.case = T) ~ 'LenderBase',
      (!is.na(Campaign.Name) & !grepl('LB|LBC', Campaign.Name, ignore.case = T)) ~ 'PBC-SBC'
      
    )
  )

karix_dump_IVR$Particulars[is.na(karix_dump_IVR$Particulars)] <- 'PBC-SBC'

karix_IVR<-karix_dump_IVR %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Karix_IVR_Sent` = sum(X15Sec.Pulse),`Karix_IVR_Cost`=Karix_IVR_Sent*0.10)

karix_DF<-karix_dump_IVR %>% group_by(Particulars) %>% dplyr::summarise(Sent = sum(X15Sec.Pulse),Cost=Sent*0.10)

KARIX_IVR<- karix_dump_IVR %>% dplyr::summarise(`Karix_IVR_Sent` = sum(X15Sec.Pulse),`Karix_IVR_Cost`=Karix_IVR_Sent*0.10)


##########################OVERALL Particulars

Overall<-list(cred_wats,cred_email,cred_sms,cred_ivr,SPivr,smsfolio,SPsmsU,SPsms,mailkoot,netcore,karix,karix_IVR) %>% reduce(full_join,by=c("Particulars","CHANNEL"))# %>% adorn_totals("row")# %>% dplyr::rename(`Spent` = 'Total')

Overall <- replace(Overall, is.na(Overall), 0)

Overall <- Overall[order(Overall$Particulars), ]

Overall_total<-Overall %>% adorn_totals("row")
Overall<-aggregate(Particulars+CHANNEL,data=Overall,FUN=sum) #%>% adorn_totals("row")%>% adorn_totals("col")

Overall$Overall_Sent = rowSums(Overall[,c("ACL_Watsapp_Sent", "ACL_Email_Sent", "ACL_SMS_Sent","ACL_IVR_Sent","Smartping_IVR_Sent","SMSFolio_SMS_Sent",
                                          "Smartping_SMS_UniqueDeli._Sent","Smartping_SMS_Sent","Mailkoot_Email_Sent","Netcore_smart_Email_Sent","Karix_SMS_Sent","Karix_IVR_Sent")])


Overall$Overall_Cost = rowSums(Overall[,c("ACL_Watsapp_Cost", "ACL_Email_Cost", "ACL_SMS_Cost","ACL_IVR_Cost","Smartping_IVR_Cost","SMSFolio_SMS_Cost",
                                          "Smartping_SMS_UniqueDeli._Cost","Smartping_SMS_Cost","Mailkoot_Email_Cost","Netcore_smart_Email_Cost","Karix_SMS_Cost","Karix_IVR_Cost")])

#######################################

#Overall_PART<-list(cred_wats_DF,cred_email_DF,cred_sms_DF,cred_ivr_DF,SPivr_DF,smsfolio_DF,SPsmsU_DF,SPsms_DF,mailkoot_DF,netcore_DF,karix_DF) %>% reduce(full_join,by=c("Particulars")) %>% group_by(Particulars)# %>% dplyr::rename(`Spent` = 'Total')

Overall_PART<-rbind(cred_wats_DF,cred_email_DF,cred_sms_DF,cred_ivr_DF,SPivr_DF,smsfolio_DF,SPsmsU_DF,SPsms_DF,mailkoot_DF,netcore_DF,karix_DF)# %>% group_by(Particulars)# %>% dplyr::rename(`Spent` = 'Total')



Overall_PART <- replace(Overall_PART, is.na(Overall_PART), 0)

Overall_PART <- Overall_PART[order(Overall_PART$Particulars), ]

library(purrr)
library(dplyr)
Overall_PART<-Overall_PART %>% group_by(Particulars)



#######################################
Overall_Panel1<-rbind(`Credence_Watsapp`= Credence_Watsapp$ACL_Watsapp_Sent, 
                      `Credence_IVR`=Credence_IVR$ACL_IVR_Sent, `Smartping_IVR`=Smartping_IVR$Smartping_IVR_Sent,
                      `Credence_SMS`=ACL_SMS_Sent,
                      `SMSFolio_SMS`=SMSFolio$SMSFolio_SMS_Sent,
                      `Smartping_SMS`=SP_SMS$SP_SMS_Sent,
                      `Credence_Email`=Credence_Email$ACL_Email_Sent, 
                      `Mailkoot_Email`=MailKoot$Mailkoot_Email_Sent,
                      `NetcoreSmart_Email`=NETCORE$Netcore_smart_Email_Sent,
                      `Karix_SMS`=KARIX$Karix_SMS_Sent,
                      `KARIX_IVR`=KARIX_IVR$Karix_IVR_Sent)
#`Smartping_SMS_UniqueDeli.`=SPSMSU$Smartping_SMS_UniqueDeli.,`Smartping_SMS`=SPSMS$Smartping_SMS)

Overall_Panel2<-rbind(`Credence_Watsapp`= Credence_Watsapp$ACL_Watsapp_Cost,
                      `Credence_IVR`=Credence_IVR$ACL_IVR_Cost,`Smartping_IVR`=Smartping_IVR$Smartping_IVR_Cost,
                      `Credence_SMS`=ACL_SMS_Cost,
                      `SMSFolio_SMS`=SMSFolio$SMSFolio_SMS_Cost,
                      `Smartping_SMS`=SP_SMS$SP_SMS_Cost,
                      `Credence_Email`=Credence_Email$ACL_Email_Cost,
                      `Mailkoot_Email`=MailKoot$Mailkoot_Email_Cost,
                      `NetcoreSmart_Email`=NETCORE$Netcore_smart_Email_Cost,
                      `Karix_SMS`=KARIX$Karix_SMS_Cost,
                      `KARIX_IVR`=KARIX_IVR$Karix_IVR_Cost)
#`Smartping_SMS_UniqueDeli.`=SPSMSU$Cost,`Smartping_SMS`=SPSMS$Cost)

Overall_Panel<-cbind(Overall_Panel1,Overall_Panel2) %>% as.data.frame()# %>% adorn_totals("row")


Overall_Panel <- replace(Overall_Panel, is.na(Overall_Panel), 0)

colnames(Overall_Panel)<-c('Sent','Cost')

Overall_Panel_FINAL<-Overall_Panel %>% adorn_totals("row") 

###############################
names(Overall)
Overall_ARS_DF<-Overall %>% select(Particulars,Overall_Sent,Overall_Cost) %>% filter(Particulars %in% c('LenderBase','PBC-SBC','RED-Profiling'))

Overall_ARS<-aggregate(cbind(Overall_Sent,Overall_Cost)~Particulars,data=Overall_ARS_DF,FUN=sum) 

Overall_ARS<-Overall_ARS %>% adorn_totals("row")#%>% adorn_totals("col")
#Overall_df <- Overall[-nrow(Overall),]

Overall_df <- Overall

Team_view <-  Overall_df %>% 
  mutate(
    Team = case_when(
      #grepl('Acko|-ISC-|APPDOWNLOAD|45-50|IDPP|RFC|Referrals|PayU|IDPP',Particulars, ignore.case = T) ~ 'Referrals',
      grepl('Acko|-ISC-',Particulars, ignore.case = T) ~ 'Acko',
      grepl('Status 45-50|45-50|IDPP',Particulars, ignore.case = T) ~ 'Status 45-50',
      grepl('PayU',Particulars, ignore.case = T) ~ 'PayU',
      Particulars %in% c('REPROFILING-RED','RED-REPRO','RED-Profiling','GREEN-Profiling','CIS-Reprofiling') ~ 'REProfiling',
      Particulars %in% c('CIS') ~ 'CIS',
      grepl('Referrals|APPDOWNLOAD|Auto_APPDOWNLOAD',Particulars, ignore.case = T) ~ 'Referrals',
      #grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      grepl('BHR-RED-DROP|BHRL-RED-DROP|BHR',Particulars, ignore.case = T) ~ 'BHR',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-RED-DROP|CHRL-RED-DROP',Particulars, ignore.case = T) ~ 'CHR-Red',
      grepl('CHR-GREEN|CHR-GREEN-Dropoff|BYS',Particulars, ignore.case = T) ~ 'CHR-Green & BYS-Amber',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC|PBC-SBC|LenderBase|LBC', Particulars, ignore.case = T) ~ 'ARS'
      
    )
  )

Team_view$Team[is.na(Team_view$Team)] <- 'ARS'

##################################
Team_view2 <-  Overall_df %>% 
  mutate(
    Team = case_when(
      grepl('Acko|-ISC-',Particulars, ignore.case = T) ~ 'Referrals',
      grepl('Auto_APPDOWNLOAD',Particulars, ignore.case = T) ~ 'Auto_APPDOWNLOAD',
      grepl('APPDOWNLOAD',Particulars, ignore.case = T) ~ 'AppDownload',
      grepl('45-50',Particulars, ignore.case = T) ~ 'Status 45-50',
      grepl('IDPP|CHR-G-IDPP',Particulars, ignore.case = T) ~ 'IDPP',
      grepl('PayU',Particulars, ignore.case = T) ~ 'PayU',
      grepl('RFC|Referrals',Particulars, ignore.case = T) ~ 'Referrals',
      grepl('CIS-REP',Particulars, ignore.case = T) ~ 'CIS-Reprofiling',
      grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      grepl('RED-Profiling',Particulars, ignore.case = T) ~ 'RED-Profiling',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-RED-DROP|CHRL-RED-DROP|CHR-GREEN|CHR-GREEN-Dropoff',Particulars, ignore.case = T) ~ 'CHR',
      grepl('BHR|BHR-RED-DROP|BHRL-RED-DROP',Particulars, ignore.case = T) ~ 'BHR',
      grepl('BYS',Particulars, ignore.case = T) ~ 'BYS-Amber',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('LenderBase|LBC', Particulars, ignore.case = T) ~ 'LenderBase',
      grepl('PBC|PBC-SBC', Particulars, ignore.case = T) ~ 'PBC',
      grepl('SBC_PBC|SBC', Particulars, ignore.case = T) ~ 'SBC'
      
    )
  )

Team_view2$Team[is.na(Team_view2$Team)] <- 'PBC'


Team_view2<-Team_view2 %>% select(Particulars,CHANNEL,Overall_Sent,Overall_Cost)#Particulars,
library(janitor)
library(purrr)
library(dplyr)
df2<-Team_view2 %>% 
  split(.[,"Particulars"]) %>% ## splits each change in cyl into a list of dataframes 
  map_df(., janitor::adorn_totals)

###########################Overall Team_COST

Overall_chk<-Overall[-c(2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22)]

Team_view3 <-  Overall_chk %>% 
  mutate(
    Team = case_when(
      grepl('Acko|-ISC-',Particulars, ignore.case = T) ~ 'Acko',
      grepl('Auto_APPDOWNLOAD',Particulars, ignore.case = T) ~ 'Auto_APPDOWNLOAD',
      grepl('APPDOWNLOAD',Particulars, ignore.case = T) ~ 'AppDownload',
      grepl('45-50',Particulars, ignore.case = T) ~ 'Status 45-50',
      grepl('IDPP|CHR-G-IDPP',Particulars, ignore.case = T) ~ 'IDPP',
      grepl('PayU',Particulars, ignore.case = T) ~ 'PayU',
      Particulars %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      Particulars %in% c('REPROFILING-GREEN|GREEN-REPRO') ~ 'GREEN-Profiling',
      Particulars %in% c('CIS-REP') ~ 'CIS-Reprofiling',
      grepl('REPROFILING-RED|TEST-REPRO|RED-REPRO', Particulars, ignore.case = T) ~ 'RED-Profiling',
      grepl('REPROFILING-GREEN|GREEN-REPRO', Particulars, ignore.case = T) ~ 'GREEN-Profiling',
      grepl('RFC|Referrals',Particulars, ignore.case = T) ~ 'Referrals',
      #grepl('CIS-REP',Particulars, ignore.case = T) ~ 'CIS-Reprofiling',
      grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      #grepl('RED-Profiling',Particulars, ignore.case = T) ~ 'RED-Profiling',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-RED-DROP|CHRL-RED-DROP|CHR-GREEN|CHR-GREEN-Dropoff',Particulars, ignore.case = T) ~ 'CHR',
      grepl('BHR|BHR-RED-DROP|BHRL-RED-DROP',Particulars, ignore.case = T) ~ 'BHR',
      grepl('BYS',Particulars, ignore.case = T) ~ 'BYS-Amber',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('LenderBase|LBC', Particulars, ignore.case = T) ~ 'LenderBase',
      grepl('PBC|PBC-SBC', Particulars, ignore.case = T) ~ 'PBC',
      grepl('SBC_PBC|SBC', Particulars, ignore.case = T) ~ 'SBC'
      
    )
  )

Team_view3$Team[is.na(Team_view3$Team)] <- 'PBC'


Team_view2<-Team_view2 %>% select(Particulars,CHANNEL,Overall_Sent,Overall_Cost)#Particulars,
library(janitor)
library(purrr)
library(dplyr)
df2<-Team_view2 %>% 
  split(.[,"Particulars"]) %>% ## splits each change in cyl into a list of dataframes 
  map_df(., janitor::adorn_totals)



names(Team_view)

Team_view_chk_1<-aggregate(cbind(Overall_Sent,Overall_Cost)~Particulars, data=Team_view,FUN=sum) #%>% adorn_totals("row")%>% adorn_totals("col")


##############################

Team_view_FINAL_DF <-  Overall_df %>% 
  mutate(
    Team = case_when(
      grepl('APPDOWNLOAD|RFC|Referrals|Auto_APPDOWNLOAD',Particulars, ignore.case = T) ~ 'Referrals',
      grepl('Acko|-ISC-',Particulars, ignore.case = T) ~ 'Insurance',
      grepl('45-50|IDPP|CHR-G-IDPP|PayU',Particulars, ignore.case = T) ~ 'Status 45-50',
      grepl('RED-Profiling|REPRO|Profiling|CIS-REP',Particulars, ignore.case = T) ~ 'REProfiling',
      grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-GREEN|CHR-GREEN-Dropoff|CHR-RED-DROP|CHRL-RED-DROP',Particulars, ignore.case = T) ~ 'CHR',
      grepl('BHR-RED-DROP|BHRL-RED-DROP|BHR',Particulars, ignore.case = T) ~ 'BHR',
      grepl('BYS',Particulars, ignore.case = T) ~ 'BYS',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC|PBC-SBC', Particulars, ignore.case = T) ~ 'ARS(PBC_SBC)',
      grepl('LenderBase|LBC', Particulars, ignore.case = T) ~ 'LenderBase',
      
    )
  )

Team_view_FINAL_DF$Team[is.na(Team_view_FINAL_DF$Team)] <- 'ARS(PBC_SBC)'



Team_view_FINAL_DF<-Team_view_FINAL_DF[-c(1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22)]
Team_view_FINAL_DF<-Team_view_FINAL_DF %>% select(Team,Overall_Sent,Overall_Cost)#Particulars,

#Team_view_FINAL_DF<-t(Team_view_FINAL_DF)

T_FINAL_DF<-aggregate(.~Team,data=Team_view_FINAL_DF,FUN=sum) %>% adorn_totals("row")
write.xlsx(T_FINAL_DF, file = "C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL_PARTICULARS.xlsx")

#######################
Team_view_FINAL_DF2 <-  Overall_df %>% 
  mutate(
    Team = case_when(
      Particulars %in% c('REPROFILING-GREEN','GREEN-REPRO','RED-Profiling','Profiling','REPRO','CIS-Reprofiling') ~ 'Product',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-GREEN|CHR-GREEN-Dropoff|CHR-RED-DROP|CHRL-RED-DROP|BHR-RED-DROP|BHRL-RED-DROP|BHR|BYS|CIS|CFP',Particulars, ignore.case = T) ~ 'Customer facing BU',
      grepl('RFC|Referrals|SBC_PBC|SBC|PBC|PBC-SBC|LenderBase|LBC',Particulars, ignore.case = T) ~ 'Lender facing BU',
      grepl('Status 45-50|45-50|IDPP|PayU|IDPP',Particulars, ignore.case = T) ~ 'Marketing',
      grepl('Acko|-ISC',Particulars, ignore.case = T) ~ 'Insurance'
      
    )
  )

Team_view_FINAL_DF2$Team[is.na(Team_view_FINAL_DF2$Team)] <- 'Lender facing BU'


Team_view_FINAL_DF2<-Team_view_FINAL_DF2[-c(1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22)]
Team_view_FINAL_DF2<-Team_view_FINAL_DF2 %>% select(Team,Overall_Sent,Overall_Cost)#Particulars,

#Team_view_FINAL_DF<-t(Team_view_FINAL_DF)

T_FINAL_DF2<-aggregate(.~Team,data=Team_view_FINAL_DF2,FUN=sum) %>% adorn_totals("row")
write.xlsx(T_FINAL_DF2, file = "C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL_SPENT.xlsx")


###########################
Team_view_FINAL <-  Overall_df %>% 
  mutate(
    Team = case_when(
      grepl('Acko|-ISC-|APPDOWNLOAD|Auto_APPDOWNLOAD|Status 45-50|IDPP|RFC|Referrals|PayU|CHR-G-IDPP|45-50',Particulars, ignore.case = T) ~ 'Referrals',
      Particulars %in% c('REPROFILING-GREEN','GREEN-REPRO','RED-Profiling','Profiling','REPRO','CIS-Reprofiling') ~ 'Product',
      grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-GREEN|CHR-GREEN-Dropoff|CHR-RED-DROP|CHRL-RED-DROP',Particulars, ignore.case = T) ~ 'CHR',
      grepl('BHR-RED-DROP|BHRL-RED-DROP|BHR',Particulars, ignore.case = T) ~ 'BHR',
      grepl('BYS',Particulars, ignore.case = T) ~ 'BYS',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC|PBC-SBC|LenderBase|LBC', Particulars, ignore.case = T) ~ 'ARS'
      
    )
  )

Team_view_FINAL$Team[is.na(Team_view_FINAL$Team)] <- 'ARS'



Team_view_FINAL<-Team_view_FINAL[-c(3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22)]
Team_view_FINAL<-Team_view_FINAL %>% select(Team,Particulars,CHANNEL,Overall_Sent,Overall_Cost)#Particulars,
library(janitor)
library(purrr)
library(dplyr)
df<-Team_view_FINAL %>% 
  split(.[,"Team"]) %>% ## splits each change in cyl into a list of dataframes 
  map_df(., janitor::adorn_totals)



ACL_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Credence=sum(ACL_Watsapp_Cost+ACL_IVR_Cost+ACL_SMS_Cost+ACL_Email_Cost))


Smartping_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Smartping_IVR=sum(Smartping_IVR_Cost))

SMSFolio_df = Team_view %>% group_by(Team) %>% dplyr::summarise(SMSFolio=sum(SMSFolio_SMS_Cost))

SP_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Smartping_SMS=sum(Smartping_SMS_Cost+Smartping_SMS_UniqueDeli._Cost))

Mailkoot_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Mailkoot=sum(Mailkoot_Email_Cost))

Netcore_smart_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Netcore_smart=sum(Netcore_smart_Email_Cost))


Karix_SMS_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Karix_SMS=sum(Karix_SMS_Cost))

Karix_IVR_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Karix_IVR=sum(Karix_IVR_Cost))

#Credence<-sum(ACL_Watsapp_Cost_df+ACL_IVR_Cost_df+ACL_SMS_Cost_df+ACL_Email_Cost_df)
df_list<-list(ACL_df,Smartping_df,SMSFolio_df,SP_df,Mailkoot_df,Netcore_smart_df,Karix_SMS_df,Karix_IVR_df)
Overall_team<-df_list %>% reduce(full_join, by='Team')

Overall_team$Overall_Cost = rowSums(Overall_team[,c("Credence", "Smartping_IVR", "SMSFolio","Smartping_SMS","Mailkoot","Netcore_smart","Karix_SMS","Karix_IVR")])

#Overall_team_df<-Overall_team %>%  filter(!row_number() %in% c(1,2,8))
Overall_team_FINAL<-Overall_team %>% adorn_totals('row')

Overall_team_df<-Overall_team

#Overall_team_df[ Overall_team_df$Team == 'CIS' , ]

total_cost<-cbind(Overall_team_df$Team,Overall_team_df$Overall_Cost) %>% as.data.frame

colnames(total_cost) <- c('Team','Cost')

total_cost<- total_cost %>% drop_na(Team)


#total_cost_df<-total_cost %>%  filter(!row_number() %in% c(1))

#View(total_cost)
############################

#ACL = data.frame(ACL_df) %>% summarise(Credence=last(ACL_df))
# 
# acl<-last(ACL_df)
# names(Overall_team)
# 
# df_cost<-aggregate(cbind(ACL_Watsapp_Cost,ACL_IVR_Cost,Smartping_IVR_Cost,ACL_Email_Cost,ACL_SMS_Cost,SMSFolio_SMS_Cost,
#                              SP_SMS_Cost,Mailkoot_Email_Cost,Netcore_smart_Email_Cost)~Team,data=Team_view,FUN=sum)
#  
# total_cost <-rowSums(df_cost[,2:ncol(df_cost)],na.rm=TRUE) %>% data.frame()
# View(total_cost)



# tot_cost<-Team_view1 %>% select(-c(ACL_Watsapp_Cost,ACL_IVR_Cost,Smartping_IVR_Cost,ACL_Email_Cost,ACL_SMS_Cost,SMSFolio_SMS_Cost,
#                                     SP_SMS_Cost,Mailkoot_Email_Cost,Netcore_smart_Email_Cost))
# 
# 
# tot_cost<-tot_cost[-c(1,7),]
# 
#team_name<-Team_view1 %>% summarise(`Unit`=Team_view1$Team)#,`COST`=Team_split$Cost)


# colnames(tot_cost)[2] <- "Cost"
# 
# View(tot_cost)
###################################

# Team_split<-aggregate(cbind(ACL_Watsapp_Cost,ACL_Email_Cost,ACL_SMS_Cost,ACL_IVR_Cost,Smartping_IVR_Cost,SMSFolio_SMS_Cost,
#                             SP_SMS_Cost,Mailkoot_Email_Cost,Netcore_smart_Email_Cost)~Team,data=Overall,FUN=sum)
# 






####################
#################################


CIS_dump <- read.csv("C:\\R\\PA_Cost\\Input\\CIS Subscribed Base.csv") %>% filter(!order_total ==0, grepl('PORTFOLIO',team,ignore.case=TRUE))

#names(CIS)
CIS_no<-nrow(CIS_dump) %>% as.data.frame()

CIS_no<-CIS_no %>% mutate(`Team` = "CIS")

colnames(CIS_no)[1] <- "Conv."


#CIS <- CIS_dump %>% dplyr::summarise(`Unit`="CIS", Conv. = count(sub_oic))


CHR_dump <- read.csv("C:\\R\\PA_Cost\\Input\\CHR-BYS Subscription Details.csv") %>% filter((!order_total %in% c(0,1415)) & grepl('PORTFOLIO',oic,ignore.case=TRUE))
#names(CIS)


CFP_dump <- read.csv("C:\\R\\PA_Cost\\Input\\CFP Subscription Details.csv") %>% filter(grepl('PORTFOLIO',channel,ignore.case=TRUE))

CFP <- CFP_dump %>% filter(channel %in% c('PORTFOLIO'))# %>%  dplyr::summarise(`Unit`="CFP", Conv. = count(oic)) 

CFP_no<-nrow(CFP) %>% as.data.frame()

CFP_no<-CFP_no %>% mutate(`Team` = "CFP")

colnames(CFP_no)[1] <- "Conv."

CHR_BYS <- CHR_dump %>% filter(!customer_type =='Red') %>% filter(!order_total %in% c(1415))# %>%  dplyr::summarise(`Unit`="CHR-GREEN & BYS-AMBER", Conv. = count(oic)) 

CHR_BYS_no<-nrow(CHR_BYS)%>% as.data.frame()


CHR_BYS_no<-CHR_BYS_no %>% mutate(`Team` = "CHR-Green & BYS-Amber")


colnames(CHR_BYS_no)[1] <- "Conv."


CHR_Red <- CHR_dump %>% filter(customer_type =='Red')# %>%  dplyr::summarise(`Unit`="CHR-RED", Conv. = count(oic)) 


CHR_Red_no<-nrow(CHR_Red) %>% as.data.frame()

CHR_Red_no<-CHR_Red_no %>% mutate(`Team` = "CHR-Red")

colnames(CHR_Red_no)[1] <- "Conv."


ref_dump <- read.xlsx("C:\\R\\PA_Cost\\Input\\Referrals's MIS.xlsx",sheet='Sheet1')

Ref = data.frame(ref_dump) %>% summarise(Conv.=last(ref_dump$PA))

Ref<-Ref %>% mutate(`Team` = "Referrals")


tot_conv<-rbind(CFP_no,CHR_BYS_no,CIS_no,CHR_Red_no,Ref)


tot_conv <- tot_conv[, c(2,1)]

tot_conv <- replace(tot_conv, is.na(tot_conv), 0)



total_cost$Conv.<-tot_conv$Conv.[match(total_cost$Team,tot_conv$Team)]


total_cost <- replace(total_cost, is.na(total_cost), 0)


#names(total_cost)
total_cost$CPL<- as.numeric(total_cost$Cost)/as.numeric(total_cost$Conv.)

#total_cost$Total = colSums(total_cost[,c("Credence", "Smartping_SMS", "SMSFolio","Smartping_IVR","Mailkoot","Netcore_smart")])

total_cost_df<-total_cost %>%  filter(!row_number() %in% c(1))

######################

#ARS cost Summary

ARS_cost_df<-Team_view %>% filter(Team %in% c('ARS'))

ARS_cost_df<-ARS_cost_df %>% select(Particulars,CHANNEL,Overall_Cost)

#Overall cost summary

Overall_particulars<-Overall %>% select(Particulars,Overall_Cost)
full_team_cost_df<-Overall_particulars %>% group_by(Particulars)


# conv_cost<-bind_cols(Overall_team,tot_conv,CPL)#conv_cost %>% mutate(`CPL`= as.numeric(unlist(tot_cost))/as.numeric(unlist(tot_conv))) %>% as.data.frame()
# 
# conv_cost <- replace(conv_cost, is.na(conv_cost), 0)
# 
# new_conv_cost <- conv_cost[-c(1:5),]
# 
# new_conv_cost<-new_conv_cost %>% adorn_totals('row')

acl_wa_df <- read.xlsx("C:\\R\\PA_Cost\\Input\\ACL_WA_Input.xlsx")

write.csv(Overall, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL-", Sys.Date()-1, '.csv'))


write.csv(Overall_Panel, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL_panel-", Sys.Date()-1, '.csv'))


write.csv(Overall_team, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL_Team-", Sys.Date()-1, '.csv'))



write.csv(total_cost, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_OVERALL_cost-", Sys.Date()-1, '.csv'))


write.csv(df, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_1stView-", Sys.Date()-1, '.csv'))


write.csv(df2, file = paste0("C:\\R\\PA_Cost\\Output\\PA_Cost_2ndView-", Sys.Date()-1, '.csv'))

write.csv(Overall_team_df, file = paste0("C:\\R\\PA_Cost\\Output\\Overall_team_df-", Sys.Date()-1, '.csv'))


#####################################PA cost MIS - mail sent


today=format(Sys.Date()-1, format="%d-%B-%Y")


#filename<-paste0("PA_Cost_OVERALL_cost-",Sys.Date()-1,".csv")

filename<-("C:\\R\\PA_Cost\\Input\\SummaryReport.csv")

myMessage = paste0("PA Cost MIS- ",today,sep='')
sender <- "referrals@creditmantri.com"
mail_LNT <-"shanthi.s@creditmantri.com"
cclist <-"shanthi.s@creditmantri.com"


# mail_LNT <- c("rakshith.thangaraj@creditmantri.com","sankarnarayanan@creditmantri.com","balamurugan.r@creditmantri.com",
#               "manikandan@creditmantri.com","shanthi.s@creditmantri.com","pradeep.e@creditmantri.com","surya.s@creditmantri.com") 
# cclist <-c("Credit-Ops@creditmantri.com")

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
    
    <h3> Overall Cost Team-wise Summary </h3>
    <p> ${print(xtable(T_FINAL_DF2, digits = 0), type = \'html\')} </p><br>
    
    <h3> Overall Cost Particulars-wise Summary </h3>
    <p> ${print(xtable(T_FINAL_DF, digits = 0), type = \'html\')} </p><br>
    
    <h3> Overall Cost Summary </h3>
    <p> ${print(xtable(total_cost, digits = 0), type = \'html\')} </p><br>
    
    <h3> Overall ARS Summary </h3>
    <p> ${print(xtable(Overall_ARS, digits = 0), type = \'html\')} </p><br>
    
    
    <h3> Overall Team-wise panel Summary </h3>
    <p> ${print(xtable(Overall_team_FINAL, digits = 0), type = \'html\')} </p><br>
    
    <h3> Overall Summary_View1 </h3>
    <p> ${print(xtable(df, digits = 0), type = \'html\')} </p><br>
    <h3> Overall Summary_View2 </h3>
    <p> ${print(xtable(df2, digits = 0), type = \'html\')} </p><br>
    
    <h3> Overall Panel-wise Summary </h3>
    <p> ${print(xtable(Overall_Panel_FINAL, digits = 0), type = \'html\')} </p><br>
    
    
    <h3> ACL WA - Traffic Analyzer </h3>
    <p> ${print(xtable(acl_wa_df, digits = 0), type = \'html\')} </p><br>
    
    
    
    
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
                                 user.name = "AKIAXJKXCWJGAXCHHIGY",
                                 passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}



