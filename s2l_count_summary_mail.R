rm(list = ls())

library(data.table)
library(readxl)
library(mailR)
library(xtable)
library(dplyr)
library(stringi)
library(purrr)
library(janitor)


today=format(Sys.Date(), format="%d-%B-%Y")



df1<-read_xlsx("C:\\R\\S2L_Manual_Lenders\\Input\\DM.xlsx") %>% select(Application_Number,Product_Name,Product_Status)

df3<-read_xlsx("C:\\R\\S2L_Manual_Lenders\\Input\\DM2.xlsx") %>% select(Application_Number,Product_Name,Product_Status)

df4<-read_xlsx("C:\\R\\S2L_Manual_Lenders\\Input\\DM3.xlsx") %>% select(Application_Number,Product_Name,Product_Status)

Final_df<-rbind(df1,df3,df4)

df2<-read_xlsx("C:\\R\\S2L_Manual_Lenders\\S2L_lender_list.xlsx")



Final_df <-  Final_df %>% 
  mutate(
    Lender = case_when(
      Product_Name == c('Axis Bank Personal Loan') ~ 'Axis Bank',
      grepl('AYE Finance', Product_Name, ignore.case = T) ~ 'AyeFinance',
      grepl('Citi Cash Back Credit Card|Citi|IndianOil Citi Platinum Card', Product_Name, ignore.case = T) ~ 'CITI',
      grepl('IIFL BUSINESS LOAN|IIFL', Product_Name, ignore.case = T) ~ 'IIFL',
      grepl('India Shelter Home Loan|India Shelter', Product_Name, ignore.case = T) ~ 'India Shelter',
      grepl('Indifi Business Loan|Indifi', Product_Name, ignore.case = T) ~ 'INDIFI',
      grepl('L&T personal Loan|LNT', Product_Name, ignore.case = T) ~ 'LNT',
      grepl('Muthoot Personal Loan', Product_Name, ignore.case = T) ~ 'Muthoot',
      grepl('PNB', Product_Name, ignore.case = T) ~ 'PNB',
      grepl('Smart Coin Short Term Loan|SmartCoin', Product_Name, ignore.case = T) ~ 'Smart Coin',
      grepl('M Capital', Product_Name, ignore.case = T) ~ 'M Capital',
      grepl('India Gold Loan', Product_Name, ignore.case = T) ~ 'India Gold',
      grepl('Rupeek', Product_Name, ignore.case = T) ~ 'Rupeek',
      grepl('Ruptok Gold Loan', Product_Name, ignore.case = T) ~ 'Ruptok',
      grepl('Bajaj', Product_Name, ignore.case = T) ~ 'Bajaj'
      
      
      
    )
  )

dfDM<- list(Final_df,df2) %>% reduce(full_join,by=c("Lender"))

dfDM <-dfDM %>%
  mutate(S2L_lenders_flag = case_when(
    grepl('Axis Bank|AyeFinance|CITI|IIFL|India Shelter|INDIFI|LNT|Muthoot|PNB|Smart Coin|M Capital|India Gold|Rupeek|Ruptok|Bajaj', Lender, ignore.case = T) ~ '1',
    is.na(Lender) ~ '0',
  )
  )

dfDM <-dfDM %>%
  mutate(New_leads_flag = case_when(
    is.na(Application_Number) ~ '0',
    !is.na(Application_Number) ~ '1'
  )
  )

dfDM <-dfDM %>%
  mutate(Remarks = case_when(
    (New_leads_flag==0) ~ 'No new lead generated',
    grepl("Axis Bank|AyeFinance|CITI|IIFL|India Shelter|LNT|Muthoot|PNB|M Capital|India Gold|Rupeek|Ruptok|Bajaj", Lender, ignore.case = T) ~ "Generated leads manually moved & sent to the lender",
    grepl('INDIFI', Lender, ignore.case = T) ~ "Generated leads manually moved & sent to BU SPOC",
    grepl('Smart Coin', Lender, ignore.case = T) ~ "Generated leads manually moved"
  )
  )

#write.csv(dfDM, paste0('C:\\R\\S2L_Manual_Lenders\\Manual_S2L-', today,'.csv'))                                     


dfDM <-  dfDM %>% filter(S2L_lenders_flag == 1) %>% group_by(Lender, Remarks) %>% dplyr::summarise(Total_Leads = n_distinct(Application_Number, na.rm = TRUE)) %>% adorn_totals("row")

#dfDM<-dfDM %>% select(dfDM,1,3,2)

dfDM<-dfDM %>% distinct(Lender, .keep_all= TRUE)

#names(dfDM)
write.csv(dfDM, paste0('C:\\R\\S2L_Manual_Lenders\\Manual_S2L.csv'))                                     


myMessage = paste0("Manual S2L MIS-",today,sep='')


sender <- "referrals@creditmantri.com"  # Replace with a valid address
recipients <- c("Ops-Referrals@creditmantri.com","referral@creditmantri.com", "r.sudarshan@creditmantri.com")  # Replace with one or more valid addresses
#filename<-paste0('C:\\R\\S2L_Manual_Lenders\\Manual_S2L-', today,'.csv')

#msg<-print(xtable(dfDM,label = message, caption = "Find below the Manual S2L Leads generated & shared to the lender"), type="html", caption.placement = "top")


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
    <h3> Overall Manual S2L Leads Summary </h3>
    <p> ${print(xtable(dfDM, digits = 0), type = \'html\')} </p><br>
    
</body>
</html>'





email <- send.mail(from = sender,
                   to = recipients,
                   subject=myMessage,
                   html = TRUE,
                   inline = T,
                   body = str_interp(msg),
                   smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                               user.name = "AKIA6IP74RHP36A2A7WY",
                               passwd = "BMRvtdvA5TlHac3vFMtO3aTFAT8wXFVEod9ZkgoftvKk" , ssl = TRUE),
                   authenticate = TRUE,
                   send = TRUE)


      
      