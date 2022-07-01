#' File to query all the data required for the MIS from the database.

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

loginfo("Referrals MIS - 1-run-queries - Initiated")

#************************************************
start_time <- Sys.time()
query = 'applications_raw.sql'

applications_raw <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/applications_raw.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE)
))%>%
  select(-applied_oic, -sku) %>%
  rename(applied_oic=attribution,
         sku = sku_name)
, error = function(e) data.frame('APPLICATION_RAW' = 'NULL')) 
         
try(logs_to_db(query, start_time))
#************************************************
start_time <- Sys.time()
query = 'aip-approved.sql'

aip_approval_raw <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/aip-approved.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)), error = function(e) data.frame('AIP_APPROVAL_RAW' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Applications Query completed")

#************************************************
start_time <- Sys.time()
query = 'fresh-profiles-inflow.sql'

fresh_inflow_raw <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/fresh-profiles-inflow.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE)
)), error = function(e) data.frame('FRESH_INFLOW_RAW' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Fresh Inflow Query completed")
#************************************************
start_time <- Sys.time()
query = 'prod-interest-funnel.sql  - MTD'

prod_int_raw_mtd <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/prod-interest-funnel.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)), error = function(e) data.frame('PROD_INT_RAW_MTD' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - product intrest Query completed")
#************************************************
start_time <- Sys.time()
query = 'prod-interest-funnel.sql - DTD'

prod_int_raw_dtd <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/prod-interest-funnel.sql"),
  list(START_DATE = DATE, END_DATE = DATE + days(1))
)), error = function(e) data.frame('PROD_INT_RAW_DTD' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Product Interest Funnels completed")
#************************************************
start_time <- Sys.time()
query = 'calls-made.sql'

calls_raw <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/calls-made.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)), error = function(e) data.frame('CALLS_RAW' = "NULL"))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Calls Query completed")
     
#************************************************
start_time <- Sys.time()
query = 'oic_form_submits.sql'

oic_form_submitts_mtd <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/oic_form_submits.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)) %>%
  left_join(select(REFERRALS_OIC, oic, name, team, oic_type)) %>%
  as.data.table(), error = function(e) data.frame('OIC_FORM_SUBMITS_MTD' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Oic Form sumbmits Query completed")
#************************************************
start_time <- Sys.time()
query = 'otp-logs.sql'

otp_logs <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, '../queries/otp-logs.sql'),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)) %>%
  left_join(select(REFERRALS_OIC, oic, name, team,oic_type)) %>%
  as.data.table(), error = function(e) data.frame('OTP_LOGS' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - OTP Query Completed")
#************************************************
start_time <- Sys.time()
query = 'fresh-inflow-funnel.sql - MTD'

fresh_inflow_funnel_mtd <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/fresh-inflow-funnel.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE)
)), error = function(e) data.frame('FRESH_INFLOW_FUNNEL_MTD' = 'NULL'))

try(logs_to_db(query, start_time))
#************************************************
start_time <- Sys.time()
query = 'fresh-inflow-funnel.sql - DTD'

fresh_inflow_funnel_dtd <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/fresh-inflow-funnel.sql"),
  list(START_DATE = DATE, END_DATE = DATE)
)), error = function(e) data.frame('FRESH_INFLOW_FUNNEL_DTD' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - Fresh Profile Queries completed")
#************************************************
start_time <- Sys.time()
query = 'referrals-cm-app-downloads.sql - MTD'

APP_DOWNLOADS_MTD <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/referrals-cm-app-downloads.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1),
       OICS = as_quoted_string(REFERRALS_OIC$oic))
)) %>%
  left_join(select(REFERRALS_OIC, oic, team,oic_type)) %>%
  mutate(download_perc = as_perc(app_downloaded / sms_sent)) %>%
  rename(Team = team, OIC = oic, `SMS Sent` = sms_sent, `Downloads` = app_downloaded,
         `Download %` = download_perc) %>%
  arrange(ifelse(OIC == 'Total', 'ZZZ', OIC)), error = function(e) data.frame('APP_DOWNLOAD_MTD' = "NULL"))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - app downloads mtd Queries completed")
#************************************************
start_time <- Sys.time()
query = 'referrals-cm-app-downloads.sql -  DTD'

APP_DOWNLOADS_DTD <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "../queries/referrals-cm-app-downloads.sql"),
  list(START_DATE = DATE, END_DATE = DATE + days(1),
       OICS = as_quoted_string(REFERRALS_OIC$oic))
)) %>%
  left_join(select(REFERRALS_OIC, oic, team,oic_type)) %>%
  select(oic, team, sms_sent, app_downloaded) %>%
  mutate(download_perc = as_perc(app_downloaded / sms_sent)) %>%
  rename(Team = team, OIC = oic, `SMS Sent` = sms_sent, `Downloads` = app_downloaded,
         `Download %` = download_perc) %>%
  arrange(ifelse(OIC == 'Total', 'ZZZ', OIC)), error = function(e) data.frame('APP_DOWNLOAD_DTD' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - app downloads dtd Queries completed")
loginfo("Referrals MIS - 1-run-queries - App Downloads Queries completed")

#************************************************
start_time <- Sys.time()
query = 'oic-login-tracking.sql'
 
oic_logins_raw <- tryCatch(data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/oic-login-tracking.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
)) %>% 
  left_join(select(REFERRALS_OIC, oic,name, team)) %>% 
  select(1:2, 14:15, 3:13), error = function(e) data.frame('OIC_LOGINS_RAW' = 'NULL'))
  
try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - oic login tracking Queries completed")
#************************************************
start_time <- Sys.time()
query = 'chr-bys-sms.sql'

chr_bys_sms <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/chr-bys-sms.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1),
       OICS = as_quoted_string(REFERRALS_OIC$oic))
) %>%
  left_join(select(REFERRALS_OIC, oic, name, team,oic_type)) %>%
  as.data.table(), error = function(e) data.frame('CHR_BYS_SMS' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - chr-bys-sms Queries completed")
#************************************************
start_time <- Sys.time()
query = 'chr-bys-payments.sql'

chr_bys_subscriptions <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/chr-bys-payments.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1),PREV_MON_START = PREV_MONTH_START,
       PREV_5_MONTH_START = PREV_5_MONTH_START,
       OICS = as_quoted_string(REFERRALS_OIC$oic))
) %>%
  left_join(select(REFERRALS_OIC, oic, name, team,oic_type)) %>%
  as.data.table(), error = function(e) data.frame(
    'user_id' = integer(),
    'payment_date' = structure(integer(), class = "Date"),
    'service_type' = character(),
    'oic' = character(),
    'team' = character())) %>% 
  as.data.table()


try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - chr-bys-payments Queries completed")
#************************************************
start_time <- Sys.time()
query = 'chr-bys-payments.sql - overall'

#' Contains STP & Assisted Conversions to be exported as a data dump.
chr_bys_subscriptions_overall <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/chr-bys-payments.sql"),
  list(START_DATE = MONTH_START,PREV_MON_START = PREV_MONTH_START,
       PREV_5_MONTH_START = PREV_5_MONTH_START,END_DATE = DATE + days(1))
) %>% {
    oic_list <- bind_rows(
      select(mutate(REFERRALS_OIC, bu = 'Referrals'), oic, team,oic_type,bu),
      select(mutate(CIS_OICS, bu = 'CIS'), oic, team, bu)
    )
    left_join(
      .,
      oic_list
    )
  } %>%
  as.data.table(), error = function(e) data.frame('CHR_BYS_SUBSCRIPTIONS_OVERALL' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - chr-bys-payments - overall Queries completed")
#************************************************

#************************************************
start_time <- Sys.time()
query = 'bhrl-payments.sql - overall'

#' Contains STP & Assisted Conversions to be exported as a data dump.
bhrl_subscriptions_overall <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/bhrl-payments.sql"),
  list(START_DATE = MONTH_START,PREV_MON_START = PREV_MONTH_START,
       PREV_5_MONTH_START = PREV_5_MONTH_START,END_DATE = DATE + days(1))
) %>% {
  oic_list <- bind_rows(
    select(mutate(REFERRALS_OIC, bu = 'Referrals'), oic, team,oic_type,bu),
    select(mutate(CIS_OICS, bu = 'CIS'), oic, team, bu)
  )
  left_join(
    .,
    oic_list
  )
} %>%
  as.data.table(), error = function(e) data.frame('BHRL_SUBSCRIPTIONS_OVERALL' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - bhrl-payments - overall Queries completed")
#************************************************

#************************************************
start_time <- Sys.time()
query = 'idpp-payments.sql - overall'

#' Contains STP & Assisted Conversions to be exported as a data dump.
idpp_subscriptions_overall <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/idpp-payments.sql"),
  list(START_DATE = MONTH_START,PREV_MON_START = PREV_MONTH_START,
       PREV_5_MONTH_START = PREV_5_MONTH_START,END_DATE = DATE + days(1))
) %>% {
  oic_list <- bind_rows(
    select(mutate(REFERRALS_OIC, bu = 'Referrals'), oic, team,oic_type,bu),
    select(mutate(CIS_OICS, bu = 'CIS'), oic, team, bu)
  )
  left_join(
    .,
    oic_list
  )
} %>%
  as.data.table(), error = function(e) data.frame('IDPP_SUBSCRIPTIONS_OVERALL' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - bhrl-payments - overall Queries completed")
#************************************************

start_time <- Sys.time()
query = 'chr-bys-interest.sql - intrest'

chr_bys_interested_base <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/chr-bys-interest.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
) %>%
  as.data.table(), error = function(e) data.frame('CHR_BYS_INTERESTED_BASE' = 'NULL'))

try(logs_to_db(query, start_time))
loginfo("Referrals MIS - 1-run-queries - chr-bys-interest - intrest Queries completed")
loginfo("Referrals MIS - 1-run-queries - OIC Login Query Queries completed")
#************************************************
start_time <- Sys.time()
query = 'chr-bys-070.sql - intrest'

chr_bys_070_base <- tryCatch(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/chr-bys-070.sql"),
  list(START_DATE = MONTH_START, END_DATE = DATE + days(1))
) %>% {
  oic_list <- bind_rows(
    select(mutate(REFERRALS_OIC, bu = 'Referrals'), oic, team,oic_type,bu),
    select(mutate(CIS_OICS, bu = 'CIS'), oic, team, bu)
  )
  left_join(
    .,
    oic_list
  )
} %>%
  as.data.table(), error = function(e) data.frame('CHR_BYS_070_BASE' = 'NULL'))

try(logs_to_db(query, start_time))

#*************************************************

OIC_LOGIN_DATA <- oic_logins_raw %>%
  rename(
    `Date` = call_date,
     OIC = oic,
    `OIC Name` = name,
    `Agency` = team,
    `Is Weekend?` = is_weekend,
    `Total Calls` = total_calls,
    `Effective Calls` = effective_calls,
    `First Call Time` = first_call_time,
    `Last Call Time` = last_call_time,
     Attendance = attendance_count,
    `Hours Logged` = hours_logged,
    `points_sku`= points_sku,
    `referrals_generated`=referrals_generated
    
  ) %>% 
  as.data.table()

ref_env <- file.path(DATA_PATH, str_interp("Referrals's MIS - ${DATE}.RData"))
save.image(ref_env)

#************************************************
#rajan changes start 29/07/2021
loginfo("Referral 90 day conversions - Query and computation Started!")

M1_START = floor_date(MONTH_START-days(2), "months")
M2_START = floor_date(M1_START-days(2), "months")
M3_START = floor_date(M2_START-days(2), "months")
M2_START_PLUS = floor_date(M2_START-days(50), "months")
months<-c(month(MONTH_START),month(M1_START),month(M2_START))
month_names<-c('January','February','March','April','May','June','July','August','September','October','November','December')

REFERRALS_conversions_90 <- 
  tryCatch(
    
    data.table(read_sql2(
      get_conn(),
      file.path(MIS_PATH, 'queries/Conversions_M2_M1_M0.sql'),
      list(
        END_DATE = DATE,
        M0=MONTH_START,
        M1=M1_START,
        M2=M2_START
      )
    )
    )
    ,error = function(e) data.frame(Metric = character(0),Total = numeric(0)))

names(REFERRALS_conversions_90)[2]<-month_names[months[1]]
names(REFERRALS_conversions_90)[6]<-month_names[months[1]]
names(REFERRALS_conversions_90)[10]<-month_names[months[1]]

names(REFERRALS_conversions_90)[3]<-month_names[months[2]]
names(REFERRALS_conversions_90)[7]<-month_names[months[2]]
names(REFERRALS_conversions_90)[11]<-month_names[months[2]]

names(REFERRALS_conversions_90)[4]<-month_names[months[3]]
names(REFERRALS_conversions_90)[8]<-month_names[months[3]]
names(REFERRALS_conversions_90)[12]<-month_names[months[3]]

REFERRALS_90_profiled<-REFERRALS_conversions_90[,1:5]
REFERRALS_90_applied<-cbind(REFERRALS_conversions_90[,1],REFERRALS_conversions_90[,6:9])
REFERRALS_90_converted<-cbind(REFERRALS_conversions_90[,1],REFERRALS_conversions_90[,10:13])

loginfo("Referral 90 day conversions - Query and computation Completed!")
#rajan changes end 29/07/2021