
library(data.table)
library(magrittr)
library(plyr)
library(dplyr)
library(lattice)
library(reshape)
library(lubridate)
library(tibble)
library(stringr)
my_cols <- c("lead_id", "customer_type", "applied_date",
             "product", "product_status", "appops_status_code", "lender",
             "profile_vintage", "sku", "attribution")
REF_MIS <- fread('C:\\R\\Referrals Dump - 2022-02-23.csv', select = my_cols)
REF_MIS = REF_MIS %>% mutate ('MTD Actual' = .N)
View (REF_MIS)
REF_TARGET <- fread('C:\\R\\Referrals_monthly_target_Feb.csv') %>%
'DRR' = (REF_TARGET$`Monthly Plan`-.N)/as.numeric(MONTH_END)-days.left,
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
  as.data.table()
