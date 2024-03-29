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


my_cols <- c("lead_id", "customer_type", "applied_date",
             "product", "product_status", "appops_status_code", "lender",
             "profile_vintage", "sku", "attribution")


#options(digits = 0)
REF_MIS <- read.xlsx("C:\\R\\07_dump.xlsx",sheet = '07_dump', detectDates = TRUE)
REF_MIS$applied_date <- convertToDate(REF_MIS$applied_date)

REF_TARGET <- fread('C:\\R\\Referrals_monthly_target_Feb.csv') %>%
  melt(id.vars = c("customer_type", "product", "SKU"),
       measure.vars = c("System", "PA", "Allset", "Regeneration", "Assisted")) %>%
  dplyr::rename ('Attribution' = variable,
         'Monthly_Plan' = value) %>% mutate('DRR' = (`Monthly_Plan`-.N)/as.numeric(MONTH_END)-days.left,
'DTD Actual' = n_distinct(lead_id[as_date(app_system$applied_date) == DATE],na.rm = T),
'MTD Actual' = .N,
'MTD Estimate' = round((.N / days.past) * days.month,0), list(customer_type, product, SKU) %>%
  .[order(customer_type, product, SKU)] %>%
  left_join(y = mtd_target, by = c('customer_type','product','SKU')) %>%
  mutate('MTD Plan' = round((`Monthly Plan`/ days.month) * days.past,0),
         'DTD %' = as_perc(`DTD Actual` / `DTD Plan`), 'MTD %' = as_perc(`MTD Actual` / `MTD Plan`), 'Backlog' = (round((`Monthly Plan`/ days.month) * days.past,0)) -.N
  ) %>% dplyr::rename('Customer Type' = customer_type, 'Product' = product) %>%
  select(
    'Customer Type','Product', 'SKU','DTD Actual', 'DTD Plan', 'DTD %',
    'MTD Actual', 'MTD Plan','MTD %','MTD Estimate', 'Monthly Plan') %>%
  as.data.table())

attributions <- c("System",'PA',"Allset","Regeneration","Assisted")
days.month <- as.numeric(DAYS_IN_MONTH)
days.past <- as.numeric(DAYS_PAST)

daysInMonth <- function(d = Sys.Date()){

  m = substr((as.character(d)), 6, 7)              # month number as string
  y = as.numeric(substr((as.character(d)), 1, 4))  # year number as numeric

  # Quick check for leap year
  leap = 0
  if ((y %% 4 == 0 & y %% 100 != 0) | y %% 400 == 0)
    leap = 1

  # Return the number of days in the month
  return(switch(m,
                '01' = 31,
                '02' = 28 + leap,  # adds 1 if leap year
                '03' = 31,
                '04' = 30,
                '05' = 31,
                '06' = 30,
                '07' = 31,
                '08' = 31,
                '09' = 30,
                '10' = 31,
                '11' = 30,
                '12' = 31))
}

DAYS_IN_MONTH = function(month, year = NULL){

  month = as.integer(month)

  if (is.null(year))
    year = as.numeric(format(Sys.Date(), '%Y'))

  dt = as.Date(paste(year, month, '01', sep = '-'))
  dates = seq(dt, by = 'month', length = 2)
  as.numeric(difftime(dates[2], dates[1], units = 'days'))
}


attributions <- c("System",'PA',"Allset","Regeneration","Assisted")
# attributions <- c("Allset")
days.month <- as.numeric(days.in.month)
days.past <- as.numeric(DAYS_PAST)

COL_START <- 3
for(team in attributions){



  mtd_target <- REF_TARGET[REF_TARGET$attribution == team,]
  mtd_target <- mtd_target[mtd_target$`Monthly Plan` != 0,]

  applications <- as.data.frame(applications)

  days.month <- case_when(team == 'Assisted' ~ 25, T ~ as.numeric(DAYS_IN_MONTH))
  days.past <- case_when(team == 'Assisted' ~ assisted_date_past[[1]], T ~ as.numeric(DAYS_PAST))
  days.left <- days.month - days.past

  days.month <- case_when(team == 'Allset' ~ 25, T ~ as.numeric(DAYS_IN_MONTH))
  days.past <- case_when(team == 'Allset' ~ assisted_date_past[[1]], T ~ as.numeric(DAYS_PAST))
  days.left <- days.month - days.past


  if(team %in% c("Allset")){

    app_system <- applications[applications$team %in% team,]


  }

  if(team %in% c("Regeneration")){

    app_system <- applications[applications$attribution %in% c('Regeneration_Others','Regeneration_PA','Regeneration_STP'),]


  }else{

    app_system <- applications[applications$attribution == team,]

  }


  #To make all varible as lower
  app_system$customer_type <- toupper(app_system$customer_type)
  app_system$product <- toupper(app_system$product)
  app_system$SKU <- toupper(app_system$SKU)
  # a[[team]]]<-app_system%>%filter (as_date(applied_date) == DATE)

  app_system <- app_system %>%
    as.data.table() %>%
    .[,.(
      # 'DRR' = (REF_TARGET$`Monthly Plan`-.N)/as.numeric(MONTH_END)-days.left,
      'DTD Actual' = n_distinct(lead_id[as_date(app_system$applied_date) == DATE],na.rm = T),
      'MTD Actual' = .N,
      'MTD Estimate' = round( (.N / days.past) * days.month),0),
      list(customer_type,product, SKU)] %>%
    .[order(customer_type,product, SKU)] %>%
    left_join(y = mtd_target, by = c('customer_type','product','SKU')) %>%
    mutate(
      'MTD Plan' = round((`Monthly Plan`/ days.month) * days.past,0),
      'DTD %' = as_perc(`DTD Actual` / `DTD Plan`),
      'MTD %' = as_perc(`MTD Actual` / `MTD Plan`)
      # 'Backlog' = (round((`Monthly Plan`/ days.month) * days.past,0)) -.N
    ) %>%
    rename(
      'Customer Type' = customer_type,
      'Product' = product) %>%
    select(
      `Customer Type`,Product, SKU,`DTD Actual`, `DTD Plan`,  `DTD %`,
      `MTD Actual`, `MTD Plan`,`MTD %`,`MTD Estimate`, `Monthly Plan`) %>%
    as.data.table()


  if(nrow(app_system)!=0){
    Backlog = as.numeric(app_system$`MTD Plan`- app_system$`MTD Actual`)
    # 'Backlog' = (round((`Monthly Plan`/ days.month) * days.past,0)) -.N,
    DRR = as.numeric(round(((app_system$`Monthly Plan`-app_system$`MTD Actual`)/as.numeric(days.month-days.left)),0))

    b1<-as.data.frame(cbind(Backlog,DRR))
    app_system<-as.data.frame(cbind(app_system,b1))
  }


  sku_list = REF_TARGET[REF_TARGET$attribution == team & REF_TARGET$`Monthly Plan` != 0,]
  skus  = paste(sku_list$customer_type, sku_list$product, sku_list$SKU)
  app_system$sku_list <- paste(app_system$`Customer Type`, app_system$Product, app_system$SKU)

  app_system <- app_system[app_system$sku_list %in% skus,]

  app_system$sku_list <- NULL


  print(app_system %>% dim())
  write_excel_table(wb, "Referrals Forecast", app_system, 3, COL_START, team, c("Product", "SKU", "Customer Type"))
  COL_START = COL_START + ncol(app_system) + 3
  column_no<-COL_START
}
