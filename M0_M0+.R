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
require(Hmisc)
library("formattable")

my_cols <- c("lead_id", "customer_type", "applied_date",
             "product", "product_status", "appops_status_code", "lender",
             "profile_vintage", "sku", "attribution")

REF_dump <- read.xlsx("C:\\R\\07_dump.xlsx",sheet = '07_dump', detectDates = TRUE)
REF_dump$applied_date <- convertToDate(REF_dump$applied_date)
#View(REF_TARGET)
REF_TARGET <- read.xlsx("C:\\R\\07_dump.xlsx",sheet = 'm0_m0+_target')

jointdataset <-merge (REF_dump, REF_TARGET, by= c('sku', 'customer_type')) %>% select("lead_id", "customer_type", "applied_date", "product_status", "appops_status_code", "lender", "profile_vintage", "sku", "attribution", "M0_STP", "M0+_STP", "M0_PA", "M0+_PA", "M0_ASSISTED", "M0+_ASSISTED", "M0_Networks", "M0_ALLSET", "M0+_ALLSET")

#attributions <- c("System",'PA',"Allset","Regeneration","Assisted")


colnames(jointdataset)

DATE <- ymd(Sys.Date() - 1)
MONTH_START <- floor_date(DATE, "months")

DAYS_IN_MONTH <- as.numeric(days_in_month(DATE))
Total_Days <- as.numeric(DAYS_IN_MONTH) - 6
DAYS_PAST <- as.numeric(day(DATE))   #w.d.finished
DAYS_LEFT <- DAYS_IN_MONTH - DAYS_PAST #Remaining days


STP_M0_DTD <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green", profile_vintage == "M0" & applied_date >= today() - days(4)) %>% dplyr::summarise('M0_Green_DTD' = n_distinct(lead_id, na.rm = TRUE))
STP_M0_MTD <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green" & profile_vintage == "M0") %>% dplyr::summarise('M0_Green_MTD'= n_distinct(lead_id, na.rm = TRUE))
STP_M0_target <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green", profile_vintage == "M0") %>% dplyr::summarise('M0_Green_Month Target' = unique(M0_STP))

STP_M0_DTD<-as.data.frame(STP_M0_DTD)
STP_M0_MTD<-as.data.frame(STP_M0_MTD)
STP_M0_target<-as.data.frame(STP_M0_target)

df<-list(STP_M0_DTD,STP_M0_MTD,STP_M0_target) %>% reduce(full_join,by=c("sku")) %>%
  mutate ('Total M0_DRR' = as.numeric(round((`STP_M0_target` - `STP_M0_MTD`)/as.numeric(DAYS_LEFT)))), 'Backlog' = as.numeric(`STP_M0_target` - `STP_M0_MTD`), 'Achieved %' = formattable::percent((`STP_M0_target` / `STP_M0_MTD`), accuracy = 0.01))

df <- replace(df, is.na(df), 0)




write_csv(df, file = "C:\\R\\Referrals_STP.csv")

