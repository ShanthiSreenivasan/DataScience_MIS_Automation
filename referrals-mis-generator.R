#' Referrals MIS
#'
#' MIS to track product-wise metrics, conversions and Assisted team calling.
 
setwd(Sys.getenv('CM_REPORTS_FOLDER'))
rm(list = ls())
MIS_NAME = 'Referral MIS' 
mis_start_time <- Sys.time()

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
library(aws.s3)
 

DATE <- ymd(Sys.Date() - 1)

args <- commandArgs(trailingOnly = T)
if (length(args) != 0 && args[1] == '--PROD') {
  TEST <- FALSE
  DATE <- ymd(Sys.Date() - 1)
} else {
  TEST <- TRUE
}

MIS_PATH <- "reports/Referrals_MIS"
DATA_PATH <- file.path(Sys.getenv('DATA_PATH'), str_interp("cm-reports-data/${DATE}"))
dir.create(DATA_PATH, recursive = T)


basicConfig()
LOGS_PATH <- file.path(Sys.getenv('LOGS_PATH'), str_interp('${DATE}'))
dir.create(LOGS_PATH, recursive = T)
addHandler(writeToFile, file = file.path(LOGS_PATH, "referrals-mis.log"))

loginfo("Referrals MIS - Initiated")
logwarn("Referrals MIS - Report Initiated")

 
source("reports/utility-functions.R")
source(file.path(MIS_PATH, "R/helpers.R"))
source("reports/excel-utilities.R")

options(error = err_func)

PREV_MONTH_DATE <- subtract_month(DATE, 1) # Same day last month
MONTH_START <- floor_date(DATE, "months")
MONTH_END <- ceiling_date(DATE, "months")-1
PREV_MONTH_START <- subtract_month(floor_date(DATE, "months"), 1)
#PREV_3_MONTH_START <- subtract_month(floor_date(DATE, "months"), 2)
PREV_5_MONTH_START <- subtract_month(floor_date(DATE, "months"), 5)
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
  .[is_active == T, .(oic, name, team,oic_type)] %>%
  distinct(oic, .keep_all = T)
CIS_OICS <- fread("config/oic_targets/CIS_OIC_MASTER_LIST.csv") %>%
  .[is_active == T & !oic %in% c('System', 'PORTFOLIO'), .(oic, name, team)] %>%
  distinct(oic, .keep_all = T)
##AGENCIES <- REFERRALS_OIC[, unique(team)]
AGENCIES <- c('Allset', 'InHouse')
 

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
## 1.1 Create Excel files   =======
#******************************************************

wb <- openxlsx::createWorkbook()
sheet_names <- c(
  "Referrals Forecast",
  "Applications Summary",
  "Lender Feedback TATs",
  "AIP Approvals",
  "Agency-wise Application Tracker",
  "Day-wise Applications",
  "Fresh Inflow Metrics",
  "OIC App Downloads",
  "OIC Logins",
  "Secondary Sales",
  "Agency-wise Applications",
  "SCC Applications",
  "07 to Bookings Funnel",
  "210 Reject Reasons",
  "SKU-wise OTP Metrics",
  "Calling Summary"
)

sheet_names <- c(sheet_names, AGENCIES)

# Create sheets
lapply(sheet_names, function(sheet) addWorksheet(wb, sheet))

output_filename <- file.path(DATA_PATH, str_interp("Referrals's MIS - ${DATE}.xlsx"))


#******************************************************
## 1.1 Run Source Files   =======
#******************************************************

source(file.path(MIS_PATH, "R/helpers.R"))
source(file.path(MIS_PATH, "R/1-run-queries.R"))
source(file.path(MIS_PATH, "R/2-clean-data.R"))
source(file.path(MIS_PATH, "R/applications-generated.R"))
source(file.path(MIS_PATH, "R/fresh-inflow-metrics.R"))
source(file.path(MIS_PATH, "R/calling-metrics.R"))
try({source(file.path(MIS_PATH, 'R/agencywise-report.R'))}, silent = T)
source(file.path(MIS_PATH, "R/applications-forecast.R"))
#******************************************************
## 1.2 Export to Excel =======
#******************************************************

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
loginfo("Referrals MIS - Excel Exported & uploaded to S3")

#******************************************************
## 2.1 Summary Tables for Email Body  =======
#******************************************************

SUMMARY_TAB <- MTD_APPLICATIONS %>%
  filter(SKU != "(all)" & `Customer Type` == "(all)") %>%
  select(-Product, -`Customer Type`) %>%
  arrange(-`(all)`)

########################################################
# Error Handling in Object
########################################################
df_name <- c("OVERALL_MTD_CALLS","SUMMARY_TAB")
for(i in df_name){if(exists(i) == FALSE){assign(i, data.frame('DATA' =  'NULL'))}}

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
                MailConfig$Product,
                'ops-Referrals@creditmantri.com',
                'marketing-ops@creditmantri.com',
                MailConfig$Finance,
                'hussain@creditmantri.com')
  
 cc_recipients <- MailConfig$Datascience

message <- str_interp("Referrals MIS - ${DATE}")

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
    <h3> Agency-wise MTD Calls </h3>
      <p> ${print(xtable(OVERALL_MTD_CALLS, digits = 0), type = \'html\')} </p><br>
    <h3> Source-wise Metrics </h3>
      <p> ${print(xtable(SUMMARY_TAB, digits = 0), type = \'html\')} </p><br>

    <h3> PROFILED 90 day METRICS </h3>
      <p> ${print(xtable(REFERRALS_90_profiled, digits = 0), type = \'html\')} </p><br>
    <h3> APPLIED 90 day METRICS </h3>
      <p> ${print(xtable(REFERRALS_90_applied, digits = 0), type = \'html\')} </p><br>
    <h3> CONVERTED90 day METRICS </h3>
      <p> ${print(xtable(REFERRALS_90_converted, digits = 0), type = \'html\')} </p><br>
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

loginfo("Referrals MIS - Email sent")
logwarn("[DONE] Referrals MIS - Report Completed")

 
##New_MIS_PATH <- "reports/Dump_generator"
##source(file.path(New_MIS_PATH,'Dump_generator.R'))
  
##New_MIS_PATH1 <- "reports/Indusind_cc_funnel/R"
##source(file.path(New_MIS_PATH1,'indusind_cc_funnel.R'))
 

New_MIS_PATH2 <- "reports"
source(file.path(New_MIS_PATH2,'indusindcc-applied-cases-dump.R'))
   
New_MIS_PATH3 <- "reports/Paysense_Funnel/R"
source(file.path(New_MIS_PATH3,'paysense_funnel.R'))
 
New_MIS_PATH4 <- "reports/icici_funnel"
source(file.path(New_MIS_PATH4,'ICICI_ptu_ptd.R'))

New_MIS_PATH5 <- "reports/IVR_MIS"
source(file.path(New_MIS_PATH5,'ivr_mis.R'))

