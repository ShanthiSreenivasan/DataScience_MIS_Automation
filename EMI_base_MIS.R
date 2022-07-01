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
library(XLConnect)


ars_dump <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\Collections Payments File.xlsx",sheet = 'Sheet1') %>% select(lead_id,EMI_cases,Date,lender,Channel,amount_paid,account_status)
names(NON_channel_mtd)
Lender_mtd<- ars_dump %>%  filter(EMI_cases %in% c(1,0)) %>% group_by(lender) %>% dplyr::summarise(`EMI Actuals Count` = n_distinct(lead_id, na.rm = TRUE), `EMI Actuals Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")# %>% mutate('DTD' = ars_dump %>% filter(EMI_cases == 1 & Date >= today() - days(3)) %>% group_by(lender) %>% dplyr::summarise('DTD' = n_distinct(lead_id, na.rm = TRUE)))

###########################EMI
emi_channel_mtd<- ars_dump %>%  filter(EMI_cases %in% c(1,0)) %>% group_by(Channel) %>% dplyr::summarise(`EMI Actuals Count` = n_distinct(lead_id, na.rm = TRUE), `EMI Actuals Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

###########################Non EMI Base updates
NON_channel_mtd<- ars_dump %>%  filter(EMI_cases == 2) %>% group_by(Channel) %>% dplyr::summarise(`Non_EMI Actuals Count` = n_distinct(lead_id, na.rm = TRUE), `Non_EMI Actuals Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")


###########################Fresh EMI cases updates

m0_channel_mtd<- ars_dump %>%  filter(EMI_cases == 1) %>% group_by(Channel) %>% dplyr::summarise(`Fresh EMI Count` = n_distinct(lead_id, na.rm = TRUE), `Fresh EMI Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

###########################Existing EMI cases updates

m1_channel_mtd<- ars_dump %>%  filter(EMI_cases == 0) %>% group_by(Channel) %>% dplyr::summarise(`Exisiting EMI Count` = n_distinct(lead_id, na.rm = TRUE), `Exisiting EMI Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

df1_MIS <- merge (emi_channel_mtd, NON_channel_mtd, by= c('Channel'))

df2_MIS <- merge (m0_channel_mtd, m1_channel_mtd, by= c('Channel'))

###########################EMI_account status-wise
emi_acc_mtd<- ars_dump %>%  filter(EMI_cases %in% c(1,0)) %>% group_by(account_status) %>% dplyr::summarise(`EMI Actuals Count` = n_distinct(lead_id, na.rm = TRUE), `EMI Actuals Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

###########################Non EMI Base updates_account status-wise
NON_acc_mtd<- ars_dump %>%  filter(EMI_cases == 2) %>% group_by(account_status) %>% dplyr::summarise(`Non_EMI Actuals Count` = n_distinct(lead_id, na.rm = TRUE), `Non_EMI Actuals Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")


###########################Fresh EMI cases updates_account status-wise

m0_acc_mtd<- ars_dump %>%  filter(EMI_cases == 1) %>% group_by(account_status) %>% dplyr::summarise(`Fresh EMI Count` = n_distinct(lead_id, na.rm = TRUE), `Fresh EMI Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

###########################Existing EMI cases updates_account status-wise

m1_acc_mtd<- ars_dump %>%  filter(EMI_cases == 0) %>% group_by(account_status) %>% dplyr::summarise(`Exisiting EMI Count` = n_distinct(lead_id, na.rm = TRUE), `Exisiting EMI Value` = sum(amount_paid, na.rm = TRUE)) %>% adorn_totals("row")

df3_MIS <- merge (emi_acc_mtd, NON_acc_mtd, by= c('account_status'))

df4_MIS <- merge (m0_acc_mtd, m1_acc_mtd, by= c('account_status'))

