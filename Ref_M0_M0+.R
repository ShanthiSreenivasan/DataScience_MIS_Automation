setwd(Sys.getenv('CM_REPORTS_FOLDER'))
rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
library(DBI)
library(RPostgreSQL)
library(RMySQL)
library(logging)
library(mailR)
library(xtable)
library(yaml)
library(jsonlite)
library(openxlsx)
library(aws.s3)
#library(tidyverse)
library(paws.storage)
library(paws)

DATE <- ymd(Sys.Date()-1)


REF_dump <- read.xlsx("C:\\R\\07_dump.xlsx",sheet = '07_dump', detectDates = TRUE)
REF_dump$applied_date <- convertToDate(REF_dump$applied_date)
#View(REF_TARGET)
REF_TARGET <- read.xlsx("C:\\R\\07_dump.xlsx",sheet = 'm0_m0+_target')

jointdataset <-merge (REF_dump, REF_TARGET, by= c('sku', 'customer_type')) %>% select("lead_id", "customer_type", "applied_date", "product_status", "appops_status_code", "lender", "profile_vintage", "sku", "attribution", "M0_STP", "M0+_STP", "M0_PA", "M0+_PA", "M0_ASSISTED", "M0+_ASSISTED", "M0_Networks", "M0_ALLSET", "M0+_ALLSET")

M0_DTD_actuals <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green", profile_vintage == "M0" & applied_date >= today() - days(4)) %>% dplyr::summarise('M0_Green_DTD' = n_distinct(lead_id, na.rm = TRUE))
M0_MTD_actuals <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green" & profile_vintage == "M0") %>% dplyr::summarise('M0_Green_MTD'= n_distinct(lead_id, na.rm = TRUE))
M0_target <- jointdataset %>% group_by(sku) %>% filter(attribution == "System", customer_type == "Green", profile_vintage == "M0") %>% dplyr::summarise('M0_Green_Month Target' = unique(M0_STP))
M0_DTD<-as.data.frame(M0_DTD_actuals)
M0_MTD<-as.data.frame(M0_MTD_actuals)
M0_target<-as.data.frame(M0_target)

STP_total_days <-as.numeric(25)
STP_off_days <-as.numeric(0)
DAYS_PAST <- as.numeric(day(DATE))
STP_remaining_days <-STP_total_days-DAYS_PAST

M0_STP_Green <- merge(M0_DTD_actuals,M0_MTD_actuals,M0_target,by.x="sku",by.y="sku",all.x=T,all.y=F) %>% reduce(full_join,by=c("sku"))

M0_STP_Green <- M0_STP_Green %>% mutate(`M0_MTD_Target`= ((`M0_target`/STP_total_days)*STP_remaining_days))


pa_target_mtd <-as.numeric(25)
pa_target_full <-STP_remaining_days

if((DAYS_PAST) <= 25){
  M0_STP_Green <- M0_STP_Green %>% mutate(`STP MTD Target`=((M0_STP_Green$M0_target)/(STP_target_mtd))*(DAYS_PAST))
}else{
  M0_STP_Green <- M0_STP_Green %>% mutate(`STP MTD Target`=((M0_STP_Green$M0_target)/(STP_target_full))*(DAYS_PAST))
}

df<-cbind(STP_M0_DTD,STP_M0_MTD,STP_M0_target) %>% mutate ('Total M0_DRR' = as.numeric(round((`STP_M0_target` - `STP_M0_MTD`)/as.numeric(DAYS_LEFT)))), 'Backlog' = as.numeric(`STP_M0_target` - `STP_M0_MTD`), 'Achieved %' = formattable::percent((`STP_M0_target` / `STP_M0_MTD`), accuracy = 0.01))
