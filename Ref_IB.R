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
ARS_IB <- read.xlsx("C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\IB MIS_Pivot.xlsx",sheet = 'MIVR_base')
names(ARS_IB)
names(ARS_IB)[6]<-c("Customer_Name")
names(ARS_IB)[8]<-c("Campaign_Name")
names(ARS_IB)[3]<-c("last_result")
#names(Output_fields)
#ARS_IB_df1 <- ARS_IB %>%
#  group_by(Campaign_Name) %>% summarise(Total_Calls = count(Campaign_Name))

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

# length(ARS_IB_df2) <- length(ARS_IB_df1)
# length(ARS_IB_df3) <- length(ARS_IB_df1)
# length(ARS_IB_df4) <- length(ARS_IB_df1)
# length(ARS_IB_df5) <- length(ARS_IB_df1)
# length(ARS_IB_df6) <- length(ARS_IB_df1)
# length(ARS_IB_df7) <- length(ARS_IB_df1)
#
# b1<-data.frame(ARS_IB_df1=ARS_IB_df1, ARS_IB_df2=ARS_IB_df2, ARS_IB_df3=ARS_IB_df3, ARS_IB_df4=ARS_IB_df4, ARS_IB_df5=ARS_IB_df5, ARS_IB_df6=ARS_IB_df6)
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

#names(final_df)
#b1 <- subset (b1, select = -c ("ARS_IB_df2.Campaign_Name"))

#test_df <- data.frame(b1)

#rm_col <- c("ARS_IB_df2.Campaign_Name", "ARS_IB_df3.Campaign_Name", "ARS_IB_df4.Campaign_Name", "ARS_IB_df5.Campaign_Name", "ARS_IB_df6.Campaign_Name", "ARS_IB_df7.Campaign_Name")
#b1 = b1[, !(colnames(b1) %in% rm_col)]

#colnames(b1)[which(names(b1) == "ARS_IB_df1.Campaign_Name")] <- "Campaign_Name"
#colnames(b1)[which(names(b1) == "ARS_IB_df1.Total_Calls")] <- "Total_Calls"
#colnames(b1)[which(names(b1) == "ARS_IB_df2.Dialled")] <- "Dialled"
#colnames(b1)[which(names(b1) == "ARS_IB_df3.Answered")] <- "Answered"
#colnames(b1)[which(names(b1) == "ARS_IB_df4.Keypress")] <- "Keypress"
#colnames(b1)[which(names(b1) == "ARS_IB_df5.Connected")] <- "Connected"
#colnames(b1)[which(names(b1) == "ARS_IB_df6.Abandoned")] <- 'Non connected / Abandoned'
#colnames(b1)[which(names(b1) == "ARS_IB_df7.Hangup")] <- "Hangup"
#b1

#b1 %>% rename(
#  Dialled = "Dialled.freq",
#  Answered = "Answered.freq"
#)
#data[!(colnames(data) %in% c('col_1','col_2'))]
#colnames(ARS_IB_df1)[which(colnames(df) == 'old_colname')] <- 'new_colname'


OB_call <- read.xlsx("C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\IB MIS_Pivot.xlsx",sheet = 'OB_call to connect_base')
colnames(OB_call)
#OB_call$addColumnDataGroups("Last.Result")
#OB_call_df1 <- OB_call %>% group_by(Last.Result) %>% dplyr::summarise(Total_Calls = n(Phone.Number.rm = TRUE))
#OB_call_df1 <- OB_call %>% group_by(Last.Result) %>% summarise(cnt = count(Last.Result)) %>% data.frame()
#OB_call_df1<- OB_call %>% group_by(Last.Result) %>% do(data.frame(nrow=nrow(.)))
#OB_call_final <- OB_call_df1 %>% adorn_totals("col")
#names(OB_call_spread)
#OB_call_df1 <- OB_call %>% group_by(Last.Result) %>% dplyr::summarise(Total_Calls = n_distinct(Phone.Number, na.rm = TRUE)) %>% adorn_totals("row") %>% dplyr::rename(`Total_Calls` = Total_Calls)
#OB_call_df1 <- OB_call %>% group_by(Last.Result) %>% dplyr::summarise(Total_Calls = n_distinct(Phone.Number, na.rm = TRUE)) %>% reshape2::dcast(. ~ Last.Result, value.var="Total_Calls")
#3.OB_call_df1 <- OB_call %>% group_by(Last.Result) %>% dplyr::summarise(Total_Calls = n_distinct(Phone.Number, na.rm = TRUE)) %>%
 #-- spread(key = Last.Result, value = Total_Calls) %>% mutate(
 #--   'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01),
 #--   'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
 #--   'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
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


#colnames(ARS_CHATBOT)
#class(ARS_CHATBOT$created_at)
ARS_CHATBOT <- read.xlsx("C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\IB MIS_Pivot.xlsx",sheet = 'ARS CHATBOT DUMP', detectDates = TRUE)# as.Date('%Y-%m-%d'))
ARS_CHATBOT$created_at <- convertToDate(ARS_CHATBOT$created_at)


#thisdate<-as.Date(format(Sys.Date()-1,'%Y-%m-%d'))

CHATBOT <- ARS_CHATBOT %>% filter((chatbot_agent_id %in% c(1, 2, 3, 4)) & (created_at >= today() - days(1))) %>% dplyr::summarise(Total_Calls = n_distinct(user_id, na.rm = TRUE)) %>% mutate(
  `Dialled` = 0, `Answered` = Total_Calls, `Keypress` = Total_Calls, `Connected` = Total_Calls, `Abandoned` = 0, 'Answered %' = percent ((`Answered` / `Total_Calls`), accuracy = 0.01), 'Key Press %' = percent ((`Keypress` / `Total_Calls`), accuracy = 0.01),
  'Connected %' = percent ((`Connected` / `Total_Calls`), accuracy = 0.01))
CHATBOT <- replace(CHATBOT, is.na(CHATBOT), 0)

#colnames(IB_Drop_dump)
IB_Drop_dump <- read.xlsx("C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\IB MIS_Pivot.xlsx",sheet = 'IB_Drop Off_pivot')
  #dplyr::summarise(Total_Calls = do(data.frame(nrow=nrow(.)) %>% mutate(
  #names(OB_call)[3]<-c("last_result")
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
final_MIS <- bind_rows(MIVR_df,OB_final_df,CHATBOT,IB_final_df) %>% adorn_totals("row")
write_csv(final_MIS, file = "C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\IB MIS_Pivot.csv")
