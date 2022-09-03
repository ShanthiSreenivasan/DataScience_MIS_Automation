rm(list=ls())
library(magrittr)
library(tibble)
library(dplyr)
library(plyr)
library(tidyr)
library(stringr)
library(lubridate)
library(openxlsx)
library(purrr)
library(janitor)
library(scales)
library(tidyverse)
library(ggrepel)
library(lookup)
library(digest)
library(data.table)
require(stats)
library(mailR)




##################################################
#View(CIS)
daily_enga_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Daily_Engagment_Stats_DUMP.csv")



Login_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Login_detail.csv")
# 
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
#       grepl('Acko|ISC',name, ignore.case = T) ~ 'Acko',
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

#str_replace_all(creden$CHANNEL, "smstrans", "Auto._SMS")

#str_replace_all(creden$CHANNEL, "whatsapp", "Auto._Watsapp")


creden<-creden %>% filter(!Particulars=='NA')
#names(creden)


#names(creden)
#View(creden)
wo_reload<-creden %>% filter(!Particulars=='Reupload',CHANNEL == "Auto._Watsapp") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_Watsapp_Sent = sum(Attempted), ACL_Watsapp_Cost=ACL_Watsapp_Sent*0.54)

cred_wats<-wo_reload

Credence_Watsapp<-creden %>% filter(!Particulars=='Reupload',CHANNEL == "Auto._Watsapp")%>% dplyr::summarise(ACL_Watsapp_Sent = sum(Attempted), ACL_Watsapp_Cost=ACL_Watsapp_Sent*0.54)

#with_reload<-creden %>% filter(CHANNEL == "whatsapp") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_Watsapp = sum(Attempted))

#reload_no<-with_reload-wo_reload
#View(with_reload)
#wo_tot<- colSums(wo_reload[, c(2:2)])


#reload_no.<- colSums(cred_wats[, c(2:2)])#, `MTD Target Achievement %` = round(`MTD Actuals (CM + Lender base)`/tar_df1)) %>% as.data.frame()

#cred_reupload<- creden %>% filter(CHANNEL == "whatsapp") %>% group_by(Particulars,CHANNEL) 

cred_email<- creden %>% filter(CHANNEL == "Auto._Email") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_Email_Sent = sum(Attempted),ACL_Email_Cost=ACL_Email_Sent*0.01)

Credence_Email<-creden %>% filter(CHANNEL == "Auto._Email") %>% dplyr::summarise(ACL_Email_Sent = sum(Attempted),ACL_Email_Cost=ACL_Email_Sent*0.01)

#cred_sms<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1,2)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms01<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms2<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(2)) %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_SMS_Sent = 2*sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

cred_sms<-rbind(cred_sms01,cred_sms2)

#Credence_SMS<- creden %>% filter(CHANNEL == "Auto._SMS") %>% dplyr::summarise(ACL_SMS_Sent = sum(Attempted),ACL_SMS_Cost=ACL_SMS_Sent*0.12)

Credence_SMS01<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(0,1)) %>% dplyr::summarise(ACL_SMS1 = sum(Attempted),Cost1=ACL_SMS1*0.12)
Credence_SMS2<- creden %>% filter(CHANNEL == "Auto._SMS", SMS_COUNT %in% c(2)) %>% dplyr::summarise(ACL_SMS2 = 2*sum(Attempted),Cost2=ACL_SMS2*0.12)


ACL_SMS_Sent<-Credence_SMS01$ACL_SMS1+Credence_SMS2$ACL_SMS2

ACL_SMS_Cost<-Credence_SMS01$Cost1+Credence_SMS2$Cost2

Credence_SMS<-cbind(ACL_SMS_Sent,ACL_SMS_Cost)

cred_ivr<- creden %>% filter(CHANNEL == "voice") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(ACL_IVR_Sent = sum(Delivered),ACL_IVR_Cost=ACL_IVR_Sent*0.22)

Credence_IVR<- creden %>% filter(CHANNEL == "voice") %>% dplyr::summarise(ACL_IVR_Sent = sum(Delivered),ACL_IVR_Cost=ACL_IVR_Sent*0.22)


#Credence_ivr<-rbind(cred_v1,cred_v2)

#Cred_Panel<-comma(Cred_Panel)
################################################SmartPing_Individual_IVR

#SPivr_dump <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP.csv") %>% mutate(`CHANNEL`="IVR")

SPivr_dump1 <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP1.csv") %>% mutate(`CHANNEL`="IVR")

SPivr_dump2 <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP2.csv") %>% mutate(`CHANNEL`="IVR")

SPivr_dump3 <- read.csv("C:\\R\\PA_Cost\\Input\\SmartPing_Individual_IVR_DUMP3.csv") %>% mutate(`CHANNEL`="IVR")


SPivr_df1<-rbind(SPivr_dump1,SPivr_dump2)


SPivr_df2<-rbind(SPivr_df1,SPivr_dump3)

SPivr_dump<-SPivr_df2

SPivr_dump<-SPivr_dump %>% filter(Overall_Call_Status == "ANSWERED")

SPivr_dump <-SPivr_dump %>%
  mutate(Duration_flag = case_when(
    Duration <= 15 ~ 1,
    Duration > 15 & Duration <= 30 ~ 2,
    Duration > 30 & Duration <= 45 ~ 3,
    Duration > 45 & Duration <= 60 ~ 4,
    Duration > 60 & Duration <= 75 ~ 5,
  ))

SPivr_dump <-  SPivr_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('SB_PB', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      
    )
  )

SPivr_dump$Particulars[is.na(SPivr_dump$Particulars)] <- 'PBC-SBC'

SPivr <- SPivr_dump

SPivr<-SPivr %>% filter(Overall_Call_Status == "ANSWERED") %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_IVR_Sent` = sum(Duration_flag),Smartping_IVR_Cost=Smartping_IVR_Sent*0.12)


Smartping_IVR<- SPivr_dump %>% dplyr::summarise(`Smartping_IVR_Sent` = sum(Duration_flag),Smartping_IVR_Cost=Smartping_IVR_Sent*0.12)


##########################SMSFolio

SMSFolio_dump <- read.csv("C:\\R\\PA_Cost\\Input\\SMSFolio_CreditMantri_Summary_DUMP.csv") %>% mutate(`CHANNEL`="SMS")

#names(SMSFolio_dump)

SMSFolio_dump <- SMSFolio_dump %>% distinct(Account, CampaignName,TotalSent, .keep_all= TRUE)


SMSFolio_dump <-  SMSFolio_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('SBC_PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      Account %in% c('creditmantri2') ~ 'Acko',
      Account %in% c('creditmantri3') & CampaignName %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      Account %in% c('creditmantri4') ~ 'LenderBase',
    )
  )

SMSFolio_dump$Particulars[is.na(SMSFolio_dump$Particulars)] <- 'PBC-SBC'



SMSFolio1 <-  SMSFolio_dump %>% filter(Account %in% c('creditmantri2','creditmantri3')) %>% 
  mutate(`Total`=TotalSent)

SMSFolio2 <-  SMSFolio_dump %>% filter(Particulars == 'LenderBase') %>% 
  mutate(`Total`=TotalSent-(UniqueSent-(UniqueDelivered+Undelivered)))


SMSFolio<-bind_rows(SMSFolio1,SMSFolio2)# %>% reduce(full_join,by=c("Particulars")) %>% adorn_totals("row")# %>% dplyr::rename(`Spent` = 'Total')

smsfolio<-SMSFolio %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`SMSFolio_SMS_Sent` = sum(Total),`SMSFolio_SMS_Cost`=SMSFolio_SMS_Sent*0.115)

SMSFolio<- SMSFolio %>% dplyr::summarise(`SMSFolio_SMS_Sent` = sum(Total),SMSFolio_SMS_Cost=SMSFolio_SMS_Sent*0.115)

################################################SmartPing_SMS


SPsms_dump <- read_xlsx("C:\\R\\PA_Cost\\Input\\Smartping_SMS_DUMP.xlsx", sheet='CampaignData') %>% mutate(`CHANNEL`="SMS")

#SPsms_login_dump <- read_xlsx("C:\\R\\PA_Cost\\Input\\SMARTPING_SMS_login_dump.xlsx")

#names(SPsms_dump)
#SPsms_dump <- SPsms_dump %>% distinct(Account, CampaignName,TotalSent, .keep_all= TRUE)

SPsms_dump$TotalSent <- as.numeric(SPsms_dump$TotalSent)

SPsms_dump$UniqueDelivered <- as.numeric(SPsms_dump$UniqueDelivered)

SPsms_dump$Undelivered <- as.numeric(SPsms_dump$Undelivered)

SPsms_dump <-  SPsms_dump %>% 
   mutate(
     Particulars = case_when(
       grepl('RFC-50-DROPOFF|RFC-50-', CampaignName, ignore.case = T) ~ '45-50',
       grepl('Acko|ISC',CampaignName, ignore.case = T) ~ 'Acko',
       grepl('www.Acko.com', URL, ignore.case = T) ~ 'Acko',
       grepl('PAYU|payu',CampaignName, ignore.case = T) ~ 'PayU',
       grepl('REPROFILING-RED',CampaignName, ignore.case = T) ~ 'RED-Profiling',
       grepl('APPDOWNLOAD|WEBLINK', CampaignName, ignore.case = T) ~ 'APPDOWNLOAD',
       grepl('CHR-RED|CHR-RENEWAL-RED', CampaignName, ignore.case = T) ~ 'CHR-RED',
       grepl('CIS|cis-test-', CampaignName, ignore.case = T) ~ 'CIS',
       grepl('RFC|rfc-test-', CampaignName, ignore.case = T) ~ 'Referrals',
       CampaignName %in% c('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST') ~ 'CHR-GREEN',
       grepl('CFP|cfp-test-|TEST-CFP-', CampaignName, ignore.case = T) ~ 'CFP',
       grepl('SBC_PBC|SBC|PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
       grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
       grepl('CHR|chr-test', CampaignName, ignore.case = T) ~ 'CHR-GREEN',
       grepl('TEST', CampaignName, ignore.case = T) ~ 'PBC-SBC',
       
     )
   )
 
 SPsms_dump$Particulars[is.na(SPsms_dump$Particulars)] <- 'PBC-SBC'

#SPsms_dump$Particulars<-SPsms_login_dump$Particulars[match(SPsms_dump$CampaignName, SPsms_login_dump$CampaignName)]

#SPsms_dump$Particulars[is.na(SPsms_dump$Particulars)] <- 'PBC-SBC'

SPsms_dump <-  SPsms_dump %>% mutate(`Total`= TotalSent - (UniqueDelivered+Undelivered))


SPsmsU<-SPsms_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_SMS_UniqueDeli._Sent` = sum(UniqueDelivered),`Smartping_SMS_UniqueDeli._Cost`=Smartping_SMS_UniqueDeli._Sent*0.105)

SPSMSU<- SPsms_dump %>% dplyr::summarise(`Smartping_SMS_UniqueDeli._Sent` = sum(UniqueDelivered),`Smartping_SMS_UniqueDeli._Cost`=Smartping_SMS_UniqueDeli._Sent*0.105)

#View(SPsms_dump)
SPsms<-SPsms_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Smartping_SMS_Sent` = sum(Total+Undelivered),`Smartping_SMS_Cost`=Smartping_SMS_Sent*0.02)

SPSMS<- SPsms_dump %>% dplyr::summarise(`Smartping_SMS_Sent` = sum(Total+Undelivered),`Smartping_SMS_Cost`=Smartping_SMS_Sent*0.02)

SP_SMS_Sent<-sum(SPSMSU$Smartping_SMS_UniqueDeli._Sent+SPSMS$Smartping_SMS_Sent)

SP_SMS_Cost<-sum(SPSMSU$Smartping_SMS_UniqueDeli._Cost+SPSMS$Smartping_SMS_Cost)

SP_SMS<-cbind(SP_SMS_Sent,SP_SMS_Cost) %>% as.data.frame()


write.xlsx(SPsms_dump,"C:\\R\\PA_Cost\\SPsms_dump.xlsx")

#write.xlsx(Overall, file = "C:\\R\\PA_Cost\\PA_Cost_OVERALL.xlsx")

################################################Mailkoot


kasp_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Kasplo_CreditMantri_Summary_DUMP.csv") %>% mutate(`CHANNEL`="email")



kasp_dump <-  kasp_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('RFC-50-DROPOFF|RFC-50-', CampaignName, ignore.case = T) ~ '45-50',
      grepl('Acko|ISC',CampaignName, ignore.case = T) ~ 'Acko',
      grepl('www.Acko.com', URL, ignore.case = T) ~ 'Acko',
      CampaignName %in% c('PAYU') ~ 'PayU',
      grepl('REPROFILING-RED|RED-REPRO', CampaignName, ignore.case = T) ~ 'RED-Profiling',
      grepl('APPDOWNLOAD', CampaignName, ignore.case = T) ~ 'APPDOWNLOAD',
      grepl('CHR-RED|CHR-RED-DROP|CHRL-RED-DROP|CHR-RENEWAL-RED',CampaignName, ignore.case = T) ~ 'CHR-RED',
      grepl('BHR-RED|BHRL-RED|BHR-RED-DROP|BHRL-RED-DROP|BHR-RENEWAL-RED',CampaignName, ignore.case = T) ~ 'BHR-RED',
      grepl('CIS|CIS-RED-DROP', CampaignName, ignore.case = T) ~ 'CIS',
      grepl('RFC', CampaignName, ignore.case = T) ~ 'Referrals',
      grepl('CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST',CampaignName, ignore.case = T) ~ 'CHR-GREEN',
      grepl('BHR-GREEN|BHR-RENEWAL-GREEN|BHR-TEST|BHRL-GREEN|BHRL-RENEWAL-GREEN|BHRL-TEST|BHR-GREEN-DROP|BHRL-GREEN-DROP',CampaignName, ignore.case = T) ~ 'BHR-GREEN',
      grepl('CFP', CampaignName, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC', CampaignName, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', CampaignName, ignore.case = T) ~ 'LenderBase',
      grepl('CHR|CHR-GREEN-DROP|CHRL-GREEN-DROP', CampaignName, ignore.case = T) ~ 'CHR-GREEN',
      grepl('TEST', CampaignName, ignore.case = T) ~ 'PBC-SBC'
      
    )
  )
kasp_dump$Particulars[is.na(kasp_dump$Particulars)] <- 'Referrals'

mailkoot<-kasp_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Mailkoot_Email_Sent` = sum(UniqueSent),`Mailkoot_Email_Cost`=Mailkoot_Email_Sent*0.009)


MailKoot<- kasp_dump %>% dplyr::summarise(`Mailkoot_Email_Sent` = sum(UniqueSent),Mailkoot_Email_Cost=Mailkoot_Email_Sent*0.009)







################################################Netcore smart


net_dump <- read.csv("C:\\R\\PA_Cost\\Input\\Email_summary_report_DUMP.csv") %>% mutate(`CHANNEL`="email")

net_dump <-  net_dump %>% 
  mutate(
    Particulars = case_when(
      grepl('RFC-50-DROPOFF|RFC-50-', Forward, ignore.case = T) ~ '45-50',
      grepl('Acko|ISC',Forward, ignore.case = T) ~ 'Acko',
      Forward %in% c('PAYU') ~ 'PayU',
      Forward %in% c('REPROFILING-RED|RED-REPRO') ~ 'RED-Profiling',
      grepl('APPDOWNLOAD|WEBLINK', Forward, ignore.case = T) ~ 'APPDOWNLOAD',
      grepl('CHR-RED|CHR-RED-DROP|CHRL-RED-DROP|CHR-RENEWAL-RED',Forward, ignore.case = T) ~ 'CHR-RED',
      grepl('BHR-RED|BHRL-RED|BHR-RED-DROP|BHRL-RED-DROP|BHR-RENEWAL-RED',Forward, ignore.case = T) ~ 'BHR-RED',
      grepl('CIS|CIS-RED-DROP', Forward, ignore.case = T) ~ 'CIS',
      grepl('RFC', Forward, ignore.case = T) ~ 'Referrals',
      grepl('BHR-GREEN-DROP|BHRL-GREEN-DROP|BHR-GREEN|BHR-RENEWAL-GREEN|BHR-TEST|BHRL-GREEN|BHRL-RENEWAL-GREEN|BHRL-TEST',Forward, ignore.case = T) ~ 'BHR-GREEN',
      grepl('CHR-GREEN-DROP|CHRL-GREEN-DROP|CHR-GREEN|CHR-RENEWAL-GREEN|CHR-TEST',Forward, ignore.case = T) ~ 'CHR-GREEN',
      grepl('CFP', Forward, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC', Forward, ignore.case = T) ~ 'PBC-SBC',
      grepl('LBC', Forward, ignore.case = T) ~ 'LenderBase',
      grepl('CHR', Forward, ignore.case = T) ~ 'CHR-GREEN',
      grepl('TEST', Forward, ignore.case = T) ~ 'PBC-SBC'
      
    )
  )

net_dump$Particulars[is.na(net_dump$Particulars)] <- 'PBC-SBC'

netcore<-net_dump %>% group_by(Particulars,CHANNEL) %>% dplyr::summarise(`Netcore_smart_Email_Sent` = sum(Total.Published),`Netcore_smart_Email_Cost`=Netcore_smart_Email_Sent*0.0125)


NETCORE<- net_dump %>% dplyr::summarise(`Netcore_smart_Email_Sent` = sum(Total.Published),Netcore_smart_Email_Cost=Netcore_smart_Email_Sent*0.0125)

#############################
##########################OVERALL Particulars

Overall<-list(cred_wats,cred_email,cred_sms,cred_ivr,SPivr,smsfolio,SPsmsU,SPsms,mailkoot,netcore) %>% reduce(full_join,by=c("Particulars","CHANNEL"))# %>% adorn_totals("row")# %>% dplyr::rename(`Spent` = 'Total')

Overall <- replace(Overall, is.na(Overall), 0)

Overall <- Overall[order(Overall$Particulars), ]

Overall$Overall_Sent = rowSums(Overall[,c("ACL_Watsapp_Sent", "ACL_Email_Sent", "ACL_SMS_Sent","ACL_IVR_Sent","Smartping_IVR_Sent","SMSFolio_SMS_Sent",
                                          "Smartping_SMS_UniqueDeli._Sent","Smartping_SMS_Sent","Mailkoot_Email_Sent","Netcore_smart_Email_Sent")])


Overall$Overall_Cost = rowSums(Overall[,c("ACL_Watsapp_Cost", "ACL_Email_Cost", "ACL_SMS_Cost","ACL_IVR_Cost","Smartping_IVR_Cost","SMSFolio_SMS_Cost",
                                          "Smartping_SMS_UniqueDeli._Cost","Smartping_SMS_Cost","Mailkoot_Email_Cost","Netcore_smart_Email_Cost")])


Overall<-Overall %>% adorn_totals("row")

#names(Overall)
# Overall_COST<-Overall %>% mutate(`Total_Sent`= sum(ACL_Watsapp_Sent,ACL_Email_Sent,ACL_SMS_Sent,ACL_IVR_Sent,Smartping_IVR_Sent,SMSFolio_SMS_Sent,Smartping_SMS_UniqueDeli._Sent,
#                                                    Smartping_SMS_Sent,Mailkoot_Email_Sent,Netcore_smart_Email_Sent), 
#                                  `Total_Cost`=sum(ACL_Watsapp_Cost,ACL_Email_Cost,ACL_SMS_Cost,ACL_IVR_Cost,Smartping_IVR_Cost,SMSFolio_SMS_Cost,
#                                                   SP_SMS_Cost,Mailkoot_Email_Cost,Netcore_smart_Email_Cost))
#######################################

Overall_df <- Overall[-nrow(Overall),]

Team_view <-  Overall_df %>% 
  mutate(
    Team = case_when(
      grepl('Acko|ISC|APPDOWNLOAD|45-50|IDPP|RFC|Referrals|PayU|CHR-G-IDPP',Particulars, ignore.case = T) ~ 'Referrals',
      grepl('CIS', Particulars, ignore.case = T) ~ 'CIS',
      grepl('CHR-RED|CHR-RED-DROP|CHRA-RED-DROP|CHR-RED-DROP|CHRL-RED-DROP|RED-Profiling|BHR-RED-DROP|BHRL-RED-DROP',Particulars, ignore.case = T) ~ 'CHR-Red',
      grepl('CHR-GREEN-Dropoff|BYS|BHR',Particulars, ignore.case = T) ~ 'CHR-Green & BYS-Amber',
      grepl('CFP', Particulars, ignore.case = T) ~ 'CFP',
      grepl('SBC_PBC|SBC|PBC|PBC-SBC|LenderBase|LBC', Particulars, ignore.case = T) ~ 'ARS',
      
    )
  )

Team_view$Team[is.na(Team_view$Team)] <- 'ARS'


#colnames(Overall)[which(names(Overall) == "ACL_Email_Cost.x")] <- "ACL_Email_Cost"

#colnames(Overall)[which(names(Overall) == "ACL_Email_Cost.y")] <- "ACL_Email_Cost"

# Team_view1<-aggregate(cbind(ACL_Watsapp_Cost,ACL_IVR_Cost,Smartping_IVR_Cost,ACL_SMS_Cost,SMSFolio_SMS_Cost,
#                             SP_SMS_Cost,ACL_Email_Cost,Mailkoot_Email_Cost,Netcore_smart_Email_Cost)~Team,data=Team_view,FUN=sum) #%>% adorn_totals("row")%>% adorn_totals("col")
# 
#Team_view_df = Team_view %>% rowwise() %>% mutate(Credence = sum(c_across(contains('ACL'))))#| contains('am')

ACL_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Credence=sum(ACL_Watsapp_Cost+ACL_IVR_Cost+ACL_SMS_Cost+ACL_Email_Cost))


Smartping_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Smartping_SMS=sum(Smartping_IVR_Cost))

SMSFolio_df = Team_view %>% group_by(Team) %>% dplyr::summarise(SMSFolio=sum(SMSFolio_SMS_Cost))

SP_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Smartping_IVR=sum(Smartping_SMS_Cost+Smartping_SMS_UniqueDeli._Cost))

Mailkoot_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Mailkoot=sum(Mailkoot_Email_Cost))

Netcore_smart_df = Team_view %>% group_by(Team) %>% dplyr::summarise(Netcore_smart=sum(Netcore_smart_Email_Cost))

#Credence<-sum(ACL_Watsapp_Cost_df+ACL_IVR_Cost_df+ACL_SMS_Cost_df+ACL_Email_Cost_df)
df_list<-list(ACL_df,Smartping_df,SMSFolio_df,SP_df,Mailkoot_df,Netcore_smart_df)
Overall_team<-df_list %>% reduce(full_join, by='Team')

Overall_team$Overall_Cost = rowSums(Overall_team[,c("Credence", "Smartping_SMS", "SMSFolio","Smartping_IVR","Mailkoot","Netcore_smart")])

total_cost<-cbind(Overall_team$Team,Overall_team$Overall_Cost) %>% as.data.frame()



colnames(total_cost) <- c('Team','Cost')

total_cost<- total_cost %>% drop_na(Team)

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


Overall_Panel1<-rbind(`Credence_Watsapp`= Credence_Watsapp$ACL_Watsapp_Sent, 
                      `Credence_IVR`=Credence_IVR$ACL_IVR_Sent, `Smartping_IVR`=Smartping_IVR$Smartping_IVR_Sent,
                      `Credence_SMS`=ACL_SMS_Sent,
                      `SMSFolio_SMS`=SMSFolio$SMSFolio_SMS_Sent,
                      `Smartping_SMS`=SP_SMS$SP_SMS_Sent,
                      `Credence_Email`=Credence_Email$ACL_Email_Sent, 
                      `Mailkoot_Email`=MailKoot$Mailkoot_Email_Sent,
                      `NetcoreSmart_Email`=NETCORE$Netcore_smart_Email_Sent)
#`Smartping_SMS_UniqueDeli.`=SPSMSU$Smartping_SMS_UniqueDeli.,`Smartping_SMS`=SPSMS$Smartping_SMS)

Overall_Panel2<-rbind(`Credence_Watsapp`= Credence_Watsapp$ACL_Watsapp_Cost,
                      `Credence_IVR`=Credence_IVR$ACL_IVR_Cost,`Smartping_IVR`=Smartping_IVR$Smartping_IVR_Cost,
                      `Credence_SMS`=ACL_SMS_Cost,
                      `SMSFolio_SMS`=SMSFolio$SMSFolio_SMS_Cost,
                      `Smartping_SMS`=SP_SMS$SP_SMS_Cost,
                      `Credence_Email`=Credence_Email$ACL_Email_Cost,
                      `Mailkoot_Email`=MailKoot$Mailkoot_Email_Cost,
                      `NetcoreSmart_Email`=NETCORE$Netcore_smart_Email_Cost)
#`Smartping_SMS_UniqueDeli.`=SPSMSU$Cost,`Smartping_SMS`=SPSMS$Cost)

Overall_Panel<-cbind(Overall_Panel1,Overall_Panel2) %>% as.data.frame()# %>% adorn_totals("row")

colnames(Overall_Panel)<-c('Sent','Cost')


#################################


CIS_dump <- read.csv("C:\\R\\PA_Cost\\Input\\CIS Subscribed Base.csv") %>% filter(!order_total ==0, grepl('PORTFOLIO',team,ignore.case=TRUE))

#names(CIS)
CIS_no<-nrow(CIS_dump) %>% as.data.frame()

CIS_no<-CIS_no %>% mutate(`Team` = "CIS")

colnames(CIS_no)[1] <- "Conv."


#CIS <- CIS_dump %>% dplyr::summarise(`Unit`="CIS", Conv. = count(sub_oic))


CHR_dump <- read.csv("C:\\R\\PA_Cost\\Input\\CHR-BYS Subscription Details.csv") %>% filter(!order_total ==0 & grepl('PORTFOLIO',oic,ignore.case=TRUE))
#names(CIS)

CFP <- CHR_dump %>% filter(order_total ==1415)# %>%  dplyr::summarise(`Unit`="CFP", Conv. = count(oic)) 

CFP_no<-nrow(CFP) %>% as.data.frame()

CFP_no<-CFP_no %>% mutate(`Team` = "CFP")

colnames(CFP_no)[1] <- "Conv."

CHR_BYS <- CHR_dump %>% filter(!order_total ==1415 & !customer_type =='Red')# %>%  dplyr::summarise(`Unit`="CHR-GREEN & BYS-AMBER", Conv. = count(oic)) 

CHR_BYS_no<-nrow(CHR_BYS)%>% as.data.frame()


CHR_BYS_no<-CHR_BYS_no %>% mutate(`Team` = "CHR-Green & BYS-Amber")


colnames(CHR_BYS_no)[1] <- "Conv."


CHR_Red <- CHR_dump %>% filter(!order_total ==1415 & customer_type =='Red')# %>%  dplyr::summarise(`Unit`="CHR-RED", Conv. = count(oic)) 


CHR_Red_no<-nrow(CHR_Red) %>% as.data.frame()

CHR_Red_no<-CHR_Red_no %>% mutate(`Team` = "CHR-Red")

colnames(CHR_Red_no)[1] <- "Conv."


ref_dump <- read_xlsx("C:\\R\\PA_Cost\\Input\\Referrals's MIS.xlsx",sheet='Sheet1')

Ref = data.frame(ref_dump) %>% summarise(Conv.=last(ref_dump$PA))

Ref<-Ref %>% mutate(`Team` = "Referrals")


tot_conv<-rbind(CFP_no,CHR_BYS_no,CIS_no,CHR_Red_no,Ref)


tot_conv <- tot_conv[, c(2,1)]

tot_conv <- replace(tot_conv, is.na(tot_conv), 0)



total_cost$Conv.<-tot_conv$Conv.[match(total_cost$Team,tot_conv$Team)]


total_cost <- replace(total_cost, is.na(total_cost), 0)


#names(total_cost)
total_cost$CPL<- as.numeric(total_cost$Cost)/as.numeric(total_cost$Conv.)

total_cost$Total = colSums(total_cost[,c("Cost", "Smartping_SMS", "SMSFolio","Smartping_IVR","Mailkoot","Netcore_smart")])

# 
# conv_cost<-bind_cols(Overall_team,tot_conv,CPL)#conv_cost %>% mutate(`CPL`= as.numeric(unlist(tot_cost))/as.numeric(unlist(tot_conv))) %>% as.data.frame()
# 
# conv_cost <- replace(conv_cost, is.na(conv_cost), 0)
# 
# new_conv_cost <- conv_cost[-c(1:5),]
# 
# new_conv_cost<-new_conv_cost %>% adorn_totals('row')

acl_wa_df <- read.xlsx("C:\\R\\PA_Cost\\Input\\ACL_WA_Input.xlsx")

write.xlsx(Overall, file = "C:\\R\\PA_Cost\\PA_Cost_OVERALL.xlsx")

write.xlsx(Overall_Panel, file = "C:\\R\\PA_Cost\\PA_Cost_OVERALL_panel.xlsx")

write.xlsx(Overall_team, file = "C:\\R\\PA_Cost\\PA_Cost_OVERALL_Team.xlsx")


write.xlsx(total_cost, file = "C:\\R\\PA_Cost\\PA_Cost_OVERALL_cost.xlsx")

#####################################PA cost MIS - mail sent


today=format(Sys.Date(), format="%d-%B-%Y")


# write.table(Overall, file = paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv"), col.names=TRUE, sep=",")
# write.table(Overall_Panel, file = paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv"), col.names=FALSE, sep=",", append=TRUE)
# write.table(Overall_Cost, file = paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv"), col.names=FALSE, sep=",", append=TRUE)
# write.table(Overall_Cost_new, file = paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv"), col.names=FALSE, sep=",", append=TRUE)
# write.csv(rbind(Overall, Overall_Panel, Overall_Cost), "filename.csv")


#data.table::fwrite(Overall_Panel, file = paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv"))



#filename<-paste0("C:\\R\\PA_Cost\\PA_Cost",today,".csv")

filename<-("C:\\R\\PA_Cost\\PA_Cost24-August-2022.csv")

myMessage = paste0("PA Cost MIS- ",today,sep='')
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
    <h3> Overall Summary </h3>
    <p> ${print(xtable(Overall, digits = 0), type = \'html\')} </p><br>
    <h3> Overall Panel-wise Summary </h3>
    <p> ${print(xtable(Overall_Panel, digits = 0), type = \'html\')} </p><br>
    <h3> Overall Team-wise panel Summary </h3>
    <p> ${print(xtable(Overall_team, digits = 0), type = \'html\')} </p><br>
    <h3> Overall Cost Summary </h3>
    <p> ${print(xtable(total_cost, digits = 0), type = \'html\')} </p><br>
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
                                 user.name = "AKIA6IP74RHPWNJ7SJUL",
                                 passwd = "BPahpceHjt3FSbdHWybStClDcsZyYZYLP/JSKE50Tyd8" , ssl = TRUE),
                     authenticate = TRUE,
                     send = TRUE)
  
}



