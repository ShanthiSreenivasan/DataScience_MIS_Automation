rm(list=ls())
library(magrittr)
library(tibble)
library(dplyr)
library(plyr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
library(openxlsx)
library(purrr)
library(janitor)
library(scales)
library(tidyverse)
require(data.table)
require(reshape2)

#########################################Docs

Ref_IB <- read.xlsx("C:\\Users\\User\\Desktop\\Ref\\IB MIS_input.xlsx",sheet = 'Docs Data')
names(Ref_IB)
names(Ref_IB)[3]<-c("Last_Result")
names(Ref_IB)[4]<-c("Home_Phone")


Ref_doc_df1 <- Ref_IB %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Total_Calls` = Total_Calls)


Ref_doc_df2 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Busy", "Invalid", "No Answer", "Failed No Lines", "Canceled", "Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Dialled` = Total_Calls)

Ref_doc_df3 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Answered` = Total_Calls)

Ref_doc_df4 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Keypress` = Total_Calls)


Ref_doc_df5 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Connected` = Total_Calls)


Ref_doc_df6 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall Abandoned")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Abandoned` = Total_Calls)

Ref_doc_df7 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Hangup` = Total_Calls)

Ref_doc_df1<-as.data.frame(Ref_doc_df1)
Ref_doc_df2<-as.data.frame(Ref_doc_df2)
Ref_doc_df3<-as.data.frame(Ref_doc_df3)
Ref_doc_df4<-as.data.frame(Ref_doc_df4)
Ref_doc_df5<-as.data.frame(Ref_doc_df5)
Ref_doc_df6<-as.data.frame(Ref_doc_df6)
Ref_doc_df7<-as.data.frame(Ref_doc_df7)

Ref_dump <- read.xlsx("C:\\Users\\User\\Desktop\\Ref\\IB MIS_input.xlsx",sheet = '07 DUMP')
doc_490 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall Abandoned")) %>% select(Home_Phone)
ref_490 <- Ref_dump %>%  filter(appops_status_code == 490) %>% select(phone_home)# %>% dplyr::summarise(Total_Calls = n_distinct(phone_home, na.rm = TRUE))

#doc_490['Home_Phone'].isin(ref_490['phone_home']).value_counts()

s = isin(doc_490$Home_Phone,ref_490$phone_home)

count_490 = doc_490['Home_Phone'] = ref_490['phone_home'] %>% dplyr::summarise(Total_Calls = n_distinct(phone_home, na.rm = TRUE)) %>% dplyr::rename(`490` = Total_Calls)

doc_990 <- Ref_IB %>%  filter(Last_Result %in% c("Answered Linkcall Abandoned")) %>% select(Home_Phone)
ref_990 <- Ref_dump %>%  filter(appops_status_code == 990) %>% select(phone_home)# %>% dplyr::summarise(Total_Calls = n_distinct(phone_home, na.rm = TRUE))
count_990 = doc_990['Home_Phone'] = ref_990['phone_home'] %>% dplyr::summarise(Total_Calls = n_distinct(phone_home, na.rm = TRUE)) %>% dplyr::rename(`990` = Total_Calls)

doc_df<- cbind('Date'=Sys.Date()-1, Ref_doc_df1,Ref_doc_df2,Ref_doc_df3,Ref_doc_df4,Ref_doc_df5,Ref_doc_df6,Ref_doc_df7,count_490,count_990) %>%
  mutate(
    'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
    'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
    'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
doc_df <- replace(doc_df, is.na(doc_df), 0)


#########################################referrals
Ref <- read.xlsx("C:\\Users\\User\\Desktop\\Ref\\IB MIS_input.xlsx",sheet = 'Referalls Data')
names(Ref)
names(Ref)[3]<-c("Last_Result")
names(Ref)[4]<-c("Home_Phone")


Ref_df1 <- Ref %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Total_Calls` = Total_Calls)


Ref_df2 <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Busy", "Invalid", "No Answer", "Failed No Lines", "Canceled", "Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Dialled` = Total_Calls)

Ref_df3 <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Answered` = Total_Calls)

Ref_df4 <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Keypress` = Total_Calls)


Ref_df5 <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Connected` = Total_Calls)


Ref_df6 <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall Abandoned")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Abandoned` = Total_Calls)

Ref_df7 <- Ref %>%  filter(Last_Result %in% c("Answered Hangup")) %>% dplyr::summarise(Total_Calls = n_distinct(Home_Phone, na.rm = TRUE)) %>% dplyr::rename(`Hangup` = Total_Calls)

Ref_df1<-as.data.frame(Ref_df1)
Ref_df2<-as.data.frame(Ref_df2)
Ref_df3<-as.data.frame(Ref_df3)
Ref_df4<-as.data.frame(Ref_df4)
Ref_df5<-as.data.frame(Ref_df5)
Ref_df6<-as.data.frame(Ref_df6)
Ref_df7<-as.data.frame(Ref_df7)

tcn_ref <- Ref %>%  filter(Last_Result %in% c("Answered Linkcall Abandoned")) %>% select(Home_Phone)
ref_count <- Ref_dump %>%  filter(attribution %in% c("Assisted")) %>% select(phone_home)
ref_leads = tcn_ref['Home_Phone'] = ref_count['phone_home'] %>% dplyr::summarise(Total_Calls = n_distinct(phone_home, na.rm = TRUE)) %>% dplyr::rename(`Referrals` = Total_Calls)

#tcn_ref['Home_Phone'].isin(ref_count['phone_home']).value_counts()
ref_df<- cbind('Date'=Sys.Date()-1,Ref_df1,Ref_df2,Ref_df3,Ref_df4,Ref_df5,Ref_df6,Ref_df7,ref_leads) %>%
  mutate(
    'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
    'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
    'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01),
    'Referrals %' = percent ((`Referrals` / `Connected`), accuracy = 0.01))
ref_df <- replace(ref_df, is.na(ref_df), 0)




# x<-seq(as.Date("2022-06-01"),as.Date("2022-06-30"),by = 1)
# x<-x[!weekdays(x) %in% c("Saturday","Sunday")]
# df1<-data.frame(x)

# y<-Sys.Date()-1
# df2<-data.frame(y)
final_MIS <- bind_rows(doc_df,ref_df) %>% adorn_totals("row")
final_MIS <- replace(final_MIS, is.na(final_MIS), 0)


write_csv(final_MIS, file = "C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\Ref IB_MIS_Pivot.csv")
