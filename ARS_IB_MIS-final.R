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

#########################################MIVR

ARS_IB <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\IB MIS_Pivot_INPUT.xlsx",sheet = 'MIVR_base')
names(ARS_IB)
names(ARS_IB)[6]<-c("Customer_Name")
names(ARS_IB)[8]<-c("Campaign_Name")
names(ARS_IB)[3]<-c("last_result")

ARS_IB_df1 <- ARS_IB %>% group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Total_Calls` = Total_Calls)


ARS_IB_df2 <- ARS_IB %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Busy", "Invalid", "No Answer", "Failed No Lines", "Canceled", "Answered Hangup")) %>%
group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Dialled` = Total_Calls)

ARS_IB_df3 <- ARS_IB %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Answered Hangup")) %>%
  group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Answered` = Total_Calls)

ARS_IB_df4 <- ARS_IB %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>%
  group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Keypress` = Total_Calls)


ARS_IB_df5 <- ARS_IB %>%  filter(last_result %in% c("Answered Linkcall")) %>%
  group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Connected` = Total_Calls)


ARS_IB_df6 <- ARS_IB %>%  filter(last_result %in% c("Answered Linkcall Abandoned")) %>%
  group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Abandoned` = Total_Calls)

ARS_IB_df7 <- ARS_IB %>%  filter(last_result %in% c("Answered Hangup")) %>%
  group_by(Campaign_Name) %>% dplyr::summarise(Total_Calls = n_distinct(User_ID, na.rm = TRUE)) %>% dplyr::rename(`Hangup` = Total_Calls)

ARS_IB_df1<-as.data.frame(ARS_IB_df1)
ARS_IB_df2<-as.data.frame(ARS_IB_df2)
ARS_IB_df3<-as.data.frame(ARS_IB_df3)
ARS_IB_df4<-as.data.frame(ARS_IB_df4)
ARS_IB_df5<-as.data.frame(ARS_IB_df5)
ARS_IB_df6<-as.data.frame(ARS_IB_df6)
ARS_IB_df7<-as.data.frame(ARS_IB_df7)

MIVR_df<-list(ARS_IB_df1,ARS_IB_df2,ARS_IB_df3,ARS_IB_df4,ARS_IB_df5,ARS_IB_df6,ARS_IB_df7) %>% reduce(full_join,by=c("Campaign_Name")) %>%
  mutate(
    'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
    'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
    'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
MIVR_df <- replace(MIVR_df, is.na(MIVR_df), 0)

#########################################OB_call

OB_call <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\IB MIS_Pivot_INPUT.xlsx",sheet = 'OB_call to connect_base')
colnames(OB_call)
names(OB_call)[3]<-c("last_result")
OB_call_df1 <- OB_call %>% dplyr::summarise(Total_Calls = n_distinct(Phone.Number, na.rm = TRUE)) %>% dplyr::rename(`Total_Calls` = Total_Calls)


OB_call_df2 <- OB_call %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Busy", "Invalid", "No Answer", "Failed No Lines", "Canceled", "Answered Hangup")) %>% dplyr::summarise(Dialled = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Dialled` = Total_Calls)

OB_call_df3 <- OB_call %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned", "Answered Hangup")) %>% dplyr::summarise(Answered = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Answered` = Total_Calls)

OB_call_df4 <- OB_call %>%  filter(last_result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>% dplyr::summarise(Keypress = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Keypress` = Total_Calls)


OB_call_df5 <- OB_call %>%  filter(last_result %in% c("Answered Linkcall")) %>% dplyr::summarise(Connected = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Connected` = Total_Calls)


OB_call_df6 <- OB_call %>%  filter(last_result %in% c("Answered Linkcall Abandoned")) %>% dplyr::summarise(Abandoned = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Abandoned` = Total_Calls)

OB_call_df7 <- OB_call %>%  filter(last_result %in% c("Answered Hangup")) %>% dplyr::summarise(Hangup = n_distinct(Phone.Number, na.rm = TRUE))# %>% dplyr::rename(`Hangup` = Total_Calls)

OB_final_df<-cbind(OB_call_df1,OB_call_df2,OB_call_df3,OB_call_df4,OB_call_df5,OB_call_df6,OB_call_df7) %>% mutate(
  'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
  'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
  'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
OB_final_df <- replace(OB_final_df, is.na(OB_final_df), 0)

#########################################CHATBOT

ARS_CHATBOT <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\IB MIS_Pivot_INPUT.xlsx",sheet = 'ARS CHATBOT DUMP', detectDates = TRUE)# as.Date('%Y-%m-%d'))
ARS_CHATBOT$created_at <- convertToDate(ARS_CHATBOT$created_at)

CHATBOT <- ARS_CHATBOT %>% filter((chatbot_agent_id %in% c(1, 2, 3, 4)) & (created_at >= today() - days(1))) %>% dplyr::summarise(Total_Calls = n_distinct(user_id, na.rm = TRUE)) %>% mutate(
  `Dialled` = 0, `Answered` = Total_Calls, `Keypress` = Total_Calls, `Connected` = Total_Calls, `Abandoned` = 0, 'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01), 'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
  'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
CHATBOT <- replace(CHATBOT, is.na(CHATBOT), 0)

#########################################IB_Drop
IB_Drop_dump <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\IB MIS_Pivot_INPUT.xlsx",sheet = 'IB_Drop Off_pivot')

IB_Drop_df1 <- IB_Drop_dump %>% dplyr::summarise(Total_Calls = n())# %>% dplyr::rename(`Total_Calls` = Total_Calls)


Dialled <- 0

IB_Drop_df3 <- IB_Drop_dump %>%  filter(Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>% dplyr::summarise(Answered = n())# %>% dplyr::rename(`Answered` = Total_Calls)

IB_Drop_df4 <- IB_Drop_dump %>%  filter(Result %in% c("Answered Linkcall", "Answered Linkcall Abandoned")) %>% dplyr::summarise(Keypress = n())# %>% dplyr::rename(`Keypress` = Total_Calls)

IB_Drop_df5 <- IB_Drop_dump %>%  filter(Result %in% c("Answered Linkcall")) %>% dplyr::summarise(Connected = n())# %>% dplyr::rename(`Connected` = Total_Calls)


IB_Drop_df6 <- IB_Drop_dump %>%  filter(Result %in% c("Answered Linkcall Abandoned")) %>% dplyr::summarise(Abandoned = n())# %>% dplyr::rename(`Abandoned` = Total_Calls)

IB_Drop_df7 <- IB_Drop_dump %>%  filter(Result %in% c("Answered Hangup")) %>% dplyr::summarise(Hangup = n())# %>% dplyr::rename(`Hangup` = Total_Calls)

IB_final_df<-cbind(IB_Drop_df1,Dialled,IB_Drop_df3,IB_Drop_df4,IB_Drop_df5,IB_Drop_df6,IB_Drop_df7) %>% mutate(
    'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
    'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
    'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
IB_final_df <- replace(IB_final_df, is.na(IB_final_df), 0)
final_MIS <- bind_rows(MIVR_df,'Call to connect	' = OB_final_df,'ChatBot' = CHATBOT,'Drop Off (CM)' = IB_final_df) %>% adorn_totals("row")


write_csv(final_MIS, file = "C:\\Users\\User\\Desktop\\MIS\\ARS\\IB MIS_Pivot_OUTPUT.csv")
