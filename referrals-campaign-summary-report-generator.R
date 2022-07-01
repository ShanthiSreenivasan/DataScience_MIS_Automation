#' Referrals Campaign Summary Report
#'
#' Report to track campaign summary metrics from dialer bases dialed out

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

DATE <- ymd(Sys.Date()-1)


args <- commandArgs(trailingOnly = T)
if (length(args) != 0 && args[1] == '--PROD') {
  TEST <- FALSE
} else {
  TEST <- TRUE
}

MIS_PATH <- "reports/Referrals_MIS"
DATA_PATH <- file.path(Sys.getenv('DATA_PATH'), str_interp("cm-reports-data/${DATE}"))
dir.create(DATA_PATH, recursive = T)


basicConfig()
LOGS_PATH <- file.path(Sys.getenv('LOGS_PATH'), str_interp('${DATE}'))
dir.create(LOGS_PATH, recursive = T)
addHandler(writeToFile, file = file.path(LOGS_PATH, "referrals-campaign-summary-report.log"))

loginfo("Referrals Campaign Summary Report - Initiated")
logwarn("Referrals Campaign Summary Report - Report Initiated")


source("reports/utility-functions.R")
source(file.path(MIS_PATH, "R/helpers.R"))
source("reports/excel-utilities.R")

if (!TEST) {
  options(error = err_func)
}

PREV_MONTH_DATE <- subtract_month(DATE, 1) # Same day last month
MONTH_START <- floor_date(DATE, "months")
PREV_MONTH_START <- subtract_month(floor_date(DATE, "months"), 1)

DAYS_IN_MONTH <- as.numeric(days_in_month(DATE))
DAYS_PAST <- day(DATE)
DAYS_LEFT <- DAYS_IN_MONTH - DAYS_PAST

DBCONFIG <- yaml.load_file("config/dbconfig.yaml")
S3CONFIG <- yaml.load_file("config/dbconfig.yaml")$s3bucket
MailConfig <- yaml.load_file("config/mail_list.yaml")


Sys.setenv(
  "AWS_ACCESS_KEY_ID" = S3CONFIG$AWS_ACCESS_KEY_ID,
  "AWS_SECRET_ACCESS_KEY" = S3CONFIG$AWS_SECRET_ACCESS_KEY,
  "AWS_DEFAULT_REGION" = S3CONFIG$AWS_DEFAULT_REGION
)

#' S3 key to use for writing data files
REPORTS_BUCKET <- get_bucket('cm-referrals-reports')
REPORTS_KEY <- str_interp('cm-referrals-reports')
MARKETING_REPORTS_BUCKET <- get_bucket('cm-marketing-reports')
MARKETING_REPORTS_KEY <- str_interp('cm-marketing-reports')
ASSISTED_REPORTS_BUCKET <- get_bucket('cm-dialler-uploads')
ASSISTED_REPORTS_KEY <- str_interp('Referrals')
CIS_REPORTS_BUCKET <- get_bucket('cm-cis-reports')
CIS_REPORTS_KEY <- str_interp('cm-cis-reports')

PRELOGIN_UTMS <- fread("config/oic_targets/PRELOGIN_UTMS.csv")

DATE_FILTERS <- list(
  "Daily" = function(date) as_date(date) == DATE,
  "MTD" = function(date) is.within.mtd(as_date(date), DATE)
)

REFERRALS_OIC <- fread("config/oic_targets/REFERRALS_OIC_MASTER_LIST.csv") %>%
  .[is_active == T, .(oic, name, team)] %>%
  distinct(oic, .keep_all = T)
CIS_OICS <- fread("config/oic_targets/CIS_OIC_MASTER_LIST.csv") %>%
  .[is_active == T & !oic %in% c('System', 'PORTFOLIO'), .(oic, name, team)] %>%
  distinct(oic, .keep_all = T)
AGENCIES <- REFERRALS_OIC[, unique(team)]

DATE_FILTERS <- list(
  "Daily" = function(date) as_date(date) == DATE,
  "MTD" = function(date) between(as_date(date), MONTH_START, DATE)
)

#*******************************************************************************
#*******************************************************************************
## 1. MAIN =======
#*******************************************************************************
#*******************************************************************************

#******************************************************
## 1.1 Run Queries   =======
#******************************************************

get_conn <- function () {
  dbConnect(
    PostgreSQL(),
    port = DBCONFIG$warehouse$port,
    user = DBCONFIG$warehouse$username,
    password = DBCONFIG$warehouse$password,
    dbname = DBCONFIG$warehouse$db,
    host = DBCONFIG$warehouse$host
  )
}

#Refresh Meterialized View
dbGetQuery(get_conn(), 'REFRESH MATERIALIZED VIEW ll_currentmonth')

loginfo("Referrals Campaign Summary Report - Running Queries ...")
logwarn("Referrals Campaign Summary Report - Running Queries ...")

CAMPAIGN_BASE_SUMMARY_DTD <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/referral-dialer-campaign-metrics.sql"),
  list(CAMPAIGN_DATE = MONTH_START, START_DATE = DATE, END_DATE = DATE + days(1))
))

loginfo("Referrals Campaign Summary Report - CAMPAIGN_BASE_SUMMARY_DTD completed")
logwarn("Referrals Campaign Summary Report - CAMPAIGN_BASE_SUMMARY_DTD completed")

CAMPAIGN_BASE_SUMMARY_MTD <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/referral-dialer-campaign-metrics.sql"),
  list(CAMPAIGN_DATE = MONTH_START, START_DATE = MONTH_START, END_DATE = DATE + days(1))
))

loginfo("Referrals Campaign Summary Report - CAMPAIGN_BASE_SUMMARY_MTD completed")
logwarn("Referrals Campaign Summary Report - CAMPAIGN_BASE_SUMMARY_MTD completed")

RECURRING_CAMPAIGN_BASE_SUMMARY <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/referral-dialer-recurring-campaign-metrics.sql"),
  list(CAMPAIGN_DATE = DATE, START_DATE = DATE, END_DATE = DATE + days(1))
))

loginfo("Referrals Campaign Summary Report - RECURRING_CAMPAIGN_BASE_SUMMARY completed")
logwarn("Referrals Campaign Summary Report - RECURRING_CAMPAIGN_BASE_SUMMARY completed")

Live_VS_CAMPAIGN_BASE_SUMMARY <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/Campaign_versus_live.sql"),
  list(CAMPAIGN_DATE = DATE, START_DATE = DATE, END_DATE = DATE + days(1))
))

CAMPAIGN_VERSUS_LIVE <- Live_VS_CAMPAIGN_BASE_SUMMARY %>%
  transpose_df('Campaign Name') %>%
  rename(Metric = metric) %>%
  select(Metric, LIVE, everything())

loginfo("Referrals Campaign Summary Report - Live_VS_CAMPAIGN_BASE_SUMMARY completed")
logwarn("Referrals Campaign Summary Report - Live_VS_CAMPAIGN_BASE_SUMMARY completed")

OIC_CAMPAIGN_BASE <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/Oic_Campaign_summary.sql"),
  list(CAMPAIGN_DATE = DATE, START_DATE = DATE, END_DATE = DATE + days(1))
))

if (nrow(OIC_CAMPAIGN_BASE) > 0) {
  OIC_CAMPAIGN_BASE[, applied_oic := NULL]

  OIC_CAMPAIGN_BASE_SUMMARY <- OIC_CAMPAIGN_BASE %>%
    left_join(select(REFERRALS_OIC, oic, team), by = 'oic') %>%
    select(oic, team, everything(OIC_CAMPAIGN_BASE))

} else {
  OIC_CAMPAIGN_BASE_SUMMARY <- create_empty_df(c('lead_id', 'oic', 'team'))
}

loginfo("Referrals Campaign Summary Report - OIC_CAMPAIGN_BASE completed")
logwarn("Referrals Campaign Summary Report - OIC_CAMPAIGN_BASE completed")



loginfo("Referrals Campaign Summary Report - All Queries Completed")
logwarn("Referrals Campaign Summary Report - All Queries Completed")


#******************************************************
## 1.2 Create Excel files   =======
#******************************************************

wb <- createWorkbook()
sheet_names <- c(
  "Campaign Summary",
  "Recurring Campaign Summary",
  "Campaign Vs. Live",
  "OIC Campaign Summary"
)
output_filename <- file.path(DATA_PATH, str_interp("Referrals Campaign Summary Report - ${DATE}.xlsx"))
# Create sheets
lapply(sheet_names, function(sheet) addWorksheet(wb, sheet))


write_excel_table(wb, "Campaign Summary", CAMPAIGN_BASE_SUMMARY_DTD,
                  3, 3, "Campaign Bases - Daily Summary", c("Campaign Name"))


write_excel_table(wb, "Campaign Summary", CAMPAIGN_BASE_SUMMARY_MTD,
                  3 + nrow(CAMPAIGN_BASE_SUMMARY_DTD) + 3, 3,
                  "Campaign Bases - MTD Summary", c("Campaign Name"))

write_excel_table(wb, "Recurring Campaign Summary", RECURRING_CAMPAIGN_BASE_SUMMARY,
                  3, 3, "Recurring Campaign Bases - Daily", c("Campaign Name"))



write_excel_table(wb, "Campaign Vs. Live", CAMPAIGN_VERSUS_LIVE,
                  3, 3, "CAMPAIGN VERSUS LIVE - Daily", c("Metric"))


write_excel_table(wb, "OIC Campaign Summary", OIC_CAMPAIGN_BASE_SUMMARY,
                  3, 3, "OIC Campaign Summary - Daily", c("oic"))


#' Adjust column widths
lapply(sheet_names, function(sheet)
  setColWidths(wb, sheet, cols = 1:120, widths = "auto", ignoreMergedCells = T)
)


saveWorkbook(wb, output_filename, overwrite = T)
aws.s3::put_object(
  file = output_filename,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(output_filename))
)
loginfo("Referrals Campaign Summary Report - Excel Exported & uploaded to S3")

#*******************************************************************************
#*******************************************************************************
## 3. Send Email =======
#*******************************************************************************
#*******************************************************************************
sender <- MailConfig$Datascience

recipients <- c(MailConfig$Referral_BU,
                MailConfig$Assitted_channel$common,
                MailConfig$Assitted_channel$ref,
                MailConfig$Assitted_channel$tnq,
                MailConfig$Marketing$ref,
                'ops-Referrals@creditmantri.com',
                MailConfig$Product)  
  

cc_recipients <- MailConfig$Datascience


message <- str_interp("Referrals Campaign Summary Report - ${DATE+1}")

if (TEST) {
  recipients <- sender
  cc_recipients <- sender
}

email_body <- '
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
<h3> ${message} </h3>
  </body>
</html>'


retry({
  send.mail(from = sender,
            to = recipients,
            cc = cc_recipients,
            subject = message,
            html = TRUE,
            inline = T,
            body = str_interp(email_body),
            smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                        user.name = "AKIAI7T5HYFCTUZMOV3Q",
                        passwd = "AtHel2jMbKGwbGlQjalkTZxEW144VM+LmgfLpNINg07E" , ssl = TRUE),
            attach.files = c(output_filename),
            authenticate = TRUE,
            send = TRUE)
})

loginfo("Referrals Campaign Summary Report - Email sent")
logwarn("[DONE] Referrals Campaign Summary Report - Report Completed")

