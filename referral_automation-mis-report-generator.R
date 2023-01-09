#'MIS Automation for Shanthi referral Tasks
#'

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

MONTH_START <- floor_date(DATE, "months")

MIS_PATH <- "reports/referral_automation"
DATA_PATH <- file.path(Sys.getenv('DATA_PATH'), str_interp("cm-reports-data/${DATE}"))
dir.create(DATA_PATH, recursive = T)

basicConfig()
LOGS_PATH <- file.path(Sys.getenv('LOGS_PATH'), str_interp('${DATE}'))
dir.create(LOGS_PATH, recursive = T)
addHandler(writeToFile, file = file.path(LOGS_PATH, "referral-automation-mis-report.log"))

loginfo("REF AUTOMATION MIS Report - Initiated")

args <- commandArgs(trailingOnly = T)

if (length(args) != 0 && args[1] == '--PROD') {
  TEST <- FALSE
} else {
  TEST <- TRUE
}

# Helper scripts
source("reports/utility-functions.R")
source("reports/excel-utilities.R")

if (!TEST) {
  options(error = err_func)
}

PREV_MONTH_DATE <- subtract_month(DATE, 1) # Same day last month
MONTH_START <- floor_date(DATE, "months")
WEEK_START = floor_date(DATE, "week")
THREE_MONTH_DATE_START <- subtract_month(floor_date(DATE, "months"), 3)
THREE_MONTH_DATE <- DATE - days(90)


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
REPORTS_BUCKET <- get_bucket('cm-cis-reports')
REPORTS_KEY <- str_interp('cm-cis-reports')
MARKETING_REPORTS_BUCKET <- get_bucket('cm-marketing-reports')
MARKETING_REPORTS_KEY <- str_interp('cm-marketing-reports')
REFERRALS_REPORTS_BUCKET <- get_bucket('cm-referrals-reports')
REFERRALS_REPORTS_KEY <- str_interp('cm-referrals-reports')


CON <-
  dbConnect(
    PostgreSQL(),
    port = DBCONFIG$warehouse$port,
    user = DBCONFIG$warehouse$username,
    password = DBCONFIG$warehouse$password,
    dbname = DBCONFIG$warehouse$db,
    host = DBCONFIG$warehouse$host
  )

DATE_FILTERS <- list(
  "DTD" = function(date) as_date(date) == DATE,
  "MTD" = function(date) is.within.mtd(as_date(date), DATE),
  "0-30 Days" = function(date) between(as_date(date),DATE - days(30), DATE),
  "30-60 Days" = function(date) between(as_date(date),DATE - days(60), DATE - days(30)),
  "60-90 Days" = function(date) between(as_date(date),DATE - days(90), DATE - days(60))
)

#*******************************************************************************
#*******************************************************************************
## 1. Run Scripts =======
#*******************************************************************************
#*******************************************************************************

#******************************************************
## 1.1 Create Excel files   =======
#************** ****************************************

wb <- openxlsx::createWorkbook()
sheet_names <- c('REF','STUCK CASES')


lapply(sheet_names, function(sheet) addWorksheet(wb, sheet))

output_filename <- file.path(DATA_PATH, str_interp("Referrals Generated MIS-Assisted_STP_PA - ${DATE}.xlsx"))


#******************************************************
## 1.2 Compute Metrics  =======
##******************************************************

source(file.path(MIS_PATH, "R/run-queries.R"))
source(file.path(MIS_PATH, "R/stuck-cases.R"))


#******************************************************
## 1.3 Export to Excel =======
#******************************************************

#' Adjust column widths
lapply(sheet_names, function(sheet)
  setColWidths(wb, sheet, cols = 1:200, widths = "auto", ignoreMergedCells = T)
)

saveWorkbook(wb, output_filename, overwrite = T)

loginfo("REF AUTOMATION MIS Report - Excel Exported")



#*******************************************************************************
#*******************************************************************************
## 2. Send Email =======
#*******************************************************************************
#*******************************************************************************
sender <- MailConfig$Datascience
recipients <- c(MailConfig$Referral_BU,
                MailConfig$Assitted_channel$common,
                MailConfig$Assitted_channel$ref,
                MailConfig$Assitted_channel$tnq,
                MailConfig$Marketing$ref,
                MailConfig$Product,
                'ops-Referrals@creditmantri.com',
                'marketing-ops@creditmantri.com',
                MailConfig$Finance)


cc_recipients <- c(sender)


message = str_interp("Referrals Generated MIS-Assisted_STP_PA - ${DATE}")
my_files =  c(output_filename)

if (TEST) {
  recipients <- sender
  cc_recipients <- sender
  message = str_interp("Referrals Generated MIS-Assisted_STP_PA - Test ${DATE}")
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
            attach.files = my_files,
            authenticate = TRUE,
            send = TRUE)
})

loginfo(" REFERRAL GENERATED MIS Report - Email Sent")
logwarn("[DONE] REFERRAL GENERATED MIS Report - Report Completed")
