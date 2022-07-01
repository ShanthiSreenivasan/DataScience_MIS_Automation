loginfo("Referrals MIS - Agencywise report - Initiated")

AGENCY_EMAIL_LIST <- list(
  #'Karvy Lucknow' = MailConfig$Ref_Agency$karvy_lucknow,
  'Allset' = MailConfig$Ref_Agency$Allset,
  'InHouse' = MailConfig$Referral_BU
)

get_agency_bucket_name <- function(s) {
  case_when(
    grepl('imarque', s, ignore.case = T) ~ 'imarque',
    grepl('mmc', s, ignore.case = T) ~ 'mmc',
    grepl('vertex.*mumbai', s, ignore.case = T) ~ 'vertex mumbai',
    grepl('kochar', s, ignore.case = T) ~ 'kochar',
    grepl('vertex.*chennai', s, ignore.case = T) ~ 'vertex chennai',
    grepl('karvy.*lucknow', s, ignore.case = T) ~ 'karvy lucknow',
    grepl('karvy.*pune', s, ignore.case = T) ~ 'karvy pune',
    grepl('allset', s, ignore.case = T) ~ 'allset',
    grepl('InHouse', s, ignore.case = T) ~ 'cm-referrals-reports'
  )
}



map(AGENCIES, function(agency_name) {
  loginfo(str_interp('${agency_name} Agency-wise Report Started ...'))
  agency_wb <- openxlsx::createWorkbook()

  sheets <- c(
    'Funnel Summary',
    'OIC-wise Metrics',
    '210 Reject Reasons',
    'App Summary',
    'OIC Login Summary'
  )

  lapply(sheets, function(sheet) addWorksheet(agency_wb, sheet))


  #' Applications Funnel
  #' ==============================

  MTD_FUNNEL <- applications %>%
    .[applied_oic %in% REFERRALS_OIC[team == agency_name, oic]] %>%
    rollup_dt(c("product", "SKU", "customer_type", "attribution"),
              compute_app_funnel, "Total") %>%
    rename(Product = product, `Customer Type` = customer_type, OIC = attribution) %>%
    arrange(Product, SKU, `Customer Type`, OIC)
  
#**************************************************
start_time <- Sys.time()
query = 'appops_270.sql'
  
  appops_270_MTD <- data.table(read_sql2(
    get_conn(),
    file.path(MIS_PATH, "queries/appops_270.sql"),
    list(START_DATE = MONTH_START, OICS = as_quoted_string(REFERRALS_OIC[team == agency_name, oic]))
  ))
try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - appops_270  Queries completed") 
#************************************************

  MTD_FUNNEL <- left_join(
    MTD_FUNNEL,
    appops_270_MTD,
    c('Product','Customer Type'='customer_type','SKU','OIC')
  )


  write_excel_table(agency_wb, 'Funnel Summary',  MTD_FUNNEL,
                    3, 3, "MTD 07 Funnel", c("Product", "SKU", "Customer Type", "OIC"))


  #' OIC-wise Metrics
  #' ==============================
  DAILY_CALLS_TAB <- get(form_agency_table(agency_name, 'Daily', 'OIC_CALLS'))
  MTD_CALLS_TAB <- get(form_agency_table(agency_name, 'MTD', 'OIC_CALLS'))

  write_excel_table(agency_wb, 'OIC-wise Metrics', DAILY_CALLS_TAB,
                    3, 3, "OIC-wise Daily Calls")
  write_excel_table(agency_wb, 'OIC-wise Metrics', MTD_CALLS_TAB,
                    3, 3 + ncol(DAILY_CALLS_TAB) + 3, "OIC-wise MTD Calls")


  #' Reject Reasons
  #' ==============================

  START_COL <- 3

  SKUWISE_REJECT_TABLES %>%
    lapply(function(x) {
      if (all(c('Reject Reason', agency_name) %in% names(x))) {
        return(subset(x, select = c('Reject Reason', agency_name)))
      } else {
        create_empty_df(c('Reject Reason', 'Total'))
      }
    }) %>%
    map2(names(.), function(sku_df, sku) {
      write_excel_table(agency_wb, "210 Reject Reasons", sku_df, 3, START_COL,
                        str_interp('${sku} 210 Reject Reasons'), c('Reject Reason'))

      START_COL <<- START_COL + ncol(sku_df) + 3
    })

  #' App Summary
  #' ==============================
  write_excel_table(agency_wb, 'App Summary',  dplyr::filter(APP_DOWNLOADS_DTD, Team == agency_name),
                    3, 3, "App Downloads - Daily Summary", c("OIC"))
  write_excel_table(agency_wb, 'App Summary', dplyr::filter(APP_DOWNLOADS_MTD, Team == agency_name),
                    3, 3 + ncol(APP_DOWNLOADS_DTD) + 3,
                    "App Downloads - MTD Summary", c("OIC"))


  #' OIC Login
  #' ==============================
  OIC_LOGIN_DATA_DTD <- OIC_LOGIN_DATA %>%
    .[OIC %in% REFERRALS_OIC[team == agency_name, oic]] %>%
    .[Date == DATE]

  OIC_LOGIN_DATA_MTD <- OIC_LOGIN_DATA %>%
    .[OIC %in% REFERRALS_OIC[team == agency_name, oic]]

  write_excel_table(agency_wb, "OIC Login Summary", OIC_LOGIN_DATA_DTD,
                    3, 3,
                    "OIC Logins - DTD Summary")
  write_excel_table(agency_wb, "OIC Login Summary", OIC_LOGIN_DATA_MTD,
                    3, 3 + ncol(OIC_LOGIN_DATA_DTD) + 3,
                    "OIC Logins - MTD Summary")



  #' Export Excel File
  #' ==============================
  #' Adjust column widths
  lapply(sheets, function(sheet)
    setColWidths(agency_wb, sheet, cols = 1:120, widths = "auto", ignoreMergedCells = T)
  )

  agency_referrals_report_fname <- file.path(
    DATA_PATH,
    'Agency-wise Referrals Reports',
    str_interp('${agency_name} Agency Referrals MIS - ${DATE}.xlsx')
  )
  dir.create(dirname(agency_referrals_report_fname), recursive = T)

  openxlsx::saveWorkbook(agency_wb, agency_referrals_report_fname, overwrite = T)

  #' Export Agency applications dump
  #' ==============================
  agency_applications_filename <- file.path(
    DATA_PATH,
    'Agency-wise Referrals Reports',
    str_interp("${agency_name} Referrals Dump - ${DATE}.csv")
  )

  applications %>%
    select(
      lead_id, date_of_referral,
      followup_date, product_status,
      lender, appops_status_code, crm_status_text = latest_crm,
      oic = assigned_oic, applied_date, ref_oic, team,
      appointment_date, oic_notes = description,
      nth,
      lender_notes,
      lender_reject_reason
    ) %>%
    .[team == agency_name] %>%
    fwrite(agency_applications_filename, dateTimeAs = "write.csv")
  aws.s3::put_object(
    file = agency_applications_filename,
    bucket = REPORTS_BUCKET,
    object = file.path(REPORTS_KEY, basename(agency_applications_filename))
  )


  agency_zipfile <- file.path(
    DATA_PATH,
    'Agency-wise Referrals Reports',
    str_interp("${agency_name} Referrals Dump - ${DATE}.zip")
  )
  try(
  zip(zipfile = agency_zipfile, files = c(agency_applications_filename)))

  try(
  aws.s3::put_object(
    file = agency_zipfile,
    bucket = REPORTS_BUCKET,
    object = file.path(REPORTS_KEY, basename(agency_zipfile))
  ))

  try(
  aws.s3::put_object(
    file = agency_zipfile,
    bucket = ASSISTED_REPORTS_BUCKET,
    object = file.path(ASSISTED_REPORTS_KEY, get_agency_bucket_name(agency_name), basename(agency_zipfile))
  ))


  loginfo(str_interp('${agency_name} Agency-wise Report Exported'))

  #******************************************************
  ## . Send Email  =======
  #******************************************************

  SENDER <- MailConfig$Datascience

  COMMON_RECIPIENTS <- c(
    MailConfig$Assitted_channel$ref,
    SENDER
  )

  AGENCY_RECIPIENTS <- AGENCY_EMAIL_LIST[[agency_name]]

  email_subject <- str_interp("${agency_name} Referrals MIS - ${DATE}")
  
  if (TEST) {
    COMMON_RECIPIENTS <- c(
      SENDER
    )
    AGENCY_RECIPIENTS <- SENDER
    email_subject <- str_interp("{TEST} - ${agency_name} Referrals MIS - ${DATE}")
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
      <h3> ${email_subject} </h3>
    </body>
  </html>'

  retry({
    send.mail(from = SENDER,
              to = AGENCY_RECIPIENTS,
              cc = COMMON_RECIPIENTS,
              subject = email_subject,
              html = TRUE,
              inline = T,
              body = str_interp(email_body),
              smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                          user.name = "AKIAI7T5HYFCTUZMOV3Q",
                          passwd = "AtHel2jMbKGwbGlQjalkTZxEW144VM+LmgfLpNINg07E" , ssl = TRUE),
              attach.files = c(agency_referrals_report_fname),
              authenticate = TRUE,
              send = TRUE)
  })
})

loginfo("Referrals MIS - Agencywise report - Completed")
