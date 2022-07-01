
loginfo("Referrals MIS - application - generated - Initiated")

MTD_APPLICATIONS <- applications %>%
  dcast_leads("product + SKU + customer_type ~ attribution", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(Product = product, `Customer Type` = customer_type)

DAILY_APPLICATIONS <- applications %>%
  .[as_date(applied_date) == DATE] %>%
  dcast_leads("product + SKU + customer_type ~ attribution", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(Product = product, `Customer Type` = customer_type)

MTD_AIP_APPROVALS <- aip_approved_cases %>%
  dcast_leads("product + SKU + customer_type ~ attribution", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(Product = product, `Customer Type` = customer_type)

DAILY_AIP_APPROVALS <- aip_approved_cases %>%
  .[as_date(applied_date) == DATE] %>%
  dcast_leads("product + SKU + customer_type ~ attribution", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(Product = product, `Customer Type` = customer_type)

MTD_AGENCYWISE_APPLICATIONS <- applications %>%
  .[!is.na(team)] %>%
  dcast_leads("SKU + customer_type ~ team", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(`Customer Type` = customer_type)

DAILY_AGENCYWISE_APPLICATIONS <- applications %>%
  .[!is.na(team)] %>%
  .[as_date(applied_date) == DATE] %>%
  dcast_leads("SKU + customer_type ~ team", value.var = "lead_id",
              fun.aggregate = n_distinct, na.rm = T) %>%
  rename(`Customer Type` = customer_type)


compute_app_funnel <- function(data) {
  data[, .(
    `Applications` = .N,
    `240` = sum(appops_status_code == '240', na.rm = T),
    `240 %` = as_perc(
      sum(appops_status_code == '240', na.rm = T)
      / .N
    ),
    `Moved to 250` = sum(!is.na(send_to_lender_oic), na.rm = T),
    `250%` = paste(round(sum(!is.na(send_to_lender_oic), na.rm = T) / .N * 100, 1), "%"),
    `Moved to 210` = sum(!is.na(not_interested_date), na.rm = T),
    `210%` = paste(round(sum(!is.na(not_interested_date), na.rm = T) / .N * 100, 1), "%"),
    # `Moved to 270` = sum(!is.na(date_of_referral), na.rm = T),
    # `270%` = paste(round(
    #   sum(!is.na(date_of_referral), na.rm = T) / sum(!is.na(send_to_lender_date), na.rm = T) * 100,
    #   1), "%"),
    `Moved to 270` = sum(appops_status_code == '270', na.rm = T),
    `270%` = as_perc(
      sum(appops_status_code == '270', na.rm = T)
      / .N
    ),
    `Feedback Received` = sum(has_received_feedback == 1, na.rm = T),
    `Feedback Received %` = as_perc(
      sum(has_received_feedback == 1, na.rm = T) / .N
    ),
    `350` = sum(appops_status_code == '350', na.rm = T),
    `350 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '350', na.rm = T)
      / .N
    ),
    `360` = sum(appops_status_code == '360', na.rm = T),
    `360 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '360', na.rm = T)
      / .N
    ),
    `370` = sum(appops_status_code == '370', na.rm = T),
    `370 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '370', na.rm = T)
      / .N
    ),
    `380` = sum(appops_status_code == '380', na.rm = T),
    `380 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '380', na.rm = T)
      / .N
    ),
    # `Moved to 390` = sum(!is.na(aip_approval_date) | is_aip_approved == 1, na.rm = T),
    # `390%` = paste(round(
    #   sum(!is.na(aip_approval_date) | is_aip_approved == 1, na.rm = T) / .N * 100,
    #   1), "%"),
    `Moved to 390` = sum(appops_status_code == '390', na.rm = T),
    `390%` = as_perc(
      sum(appops_status_code == '390', na.rm = T)
      / .N
    ),
    #Adding 399 & 400-Nirmal
    #Changes Begin
    `399` = sum(appops_status_code == '399', na.rm = T),
    `399 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '399', na.rm = T)
      / .N
    ),
    `400` = sum(appops_status_code == '400', na.rm = T),
    `400 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '400', na.rm = T)
      / .N
    ),
    #Changes End
    `450` = sum(appops_status_code == '450', na.rm = T),
    `450 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '450', na.rm = T)
      / .N
    ),
    `460` = sum(appops_status_code == '460', na.rm = T),
    `460 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '460', na.rm = T)
      / .N
    ),
    `470` = sum(appops_status_code == '470', na.rm = T),
    `470 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '470', na.rm = T)
      / .N
    ),
    `480` = sum(appops_status_code == '480', na.rm = T),
    `480 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '480', na.rm = T)
      / .N
    ),
    `490` = sum(appops_status_code == '490', na.rm = T),
    `490 %` = as_perc(
      sum(has_received_feedback == 1 & appops_status_code == '490', na.rm = T)
      / .N
    )
  )]
}

compute_agency_tracker <- function(data) {
  data[, .(
    `Daily Actuals` = sum(DATE_FILTERS[['Daily']](applied_date), na.rm = T),
    `Daily Plan` = NA,
    `MTD Actuals` = sum(DATE_FILTERS[['MTD']](applied_date), na.rm = T),
    `MTD Plan` = NA,
    `Monthly Plan` = NA,

    `MTD 270` = sum(DATE_FILTERS[['MTD']](applied_date) & pass_270 == 1, na.rm = T),
    `MTD 270 %` = as_perc(
      sum(DATE_FILTERS[['MTD']](applied_date) & pass_270 == 1, na.rm = T) /
        sum(DATE_FILTERS[['MTD']](applied_date), na.rm = T)),


    `MTD AIP` = sum(DATE_FILTERS[['MTD']](applied_date) & !is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T),
    `MTD AIP %` = as_perc(
      sum(DATE_FILTERS[['MTD']](applied_date) & !is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T) /
        sum(DATE_FILTERS[['MTD']](applied_date), na.rm = T))
  )]
}

compute_postapp_metrics <- function(data) {
  data[, .(
    Applications = .N,
    `270` = sum(!is.na(date_of_referral), na.rm = T),
    `270 (> 7d TAT) %`= as_perc(
      sum(date_diff(date_of_referral, applied_date) > 7, na.rm = T) /
        sum(!is.na(date_of_referral), na.rm = T)
    ),
    `Feedback Received` = sum(has_received_feedback == 1, na.rm = T),
    `Feedback Received %` = as_perc(
      sum(has_received_feedback == 1, na.rm = T) / sum(!is.na(date_of_referral), na.rm = T)
    ),
    `Feedback Received (> 7d TAT) %`= as_perc(
      sum(date_diff(feedback_received_date, date_of_referral) > 7, na.rm = T) /
        sum(has_received_feedback == 1, na.rm = T)
    ),
    `300` = sum(pass_300 == 1, na.rm = T),
    `390` = sum(!is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T),
    `390 (> 7d TAT) %` = as_perc(
      sum(date_diff(aip_approval_date, date_of_referral) > 7 & is_aip_approved == 1, na.rm = T) /
        sum(!is.na(date_of_referral), na.rm = T)
    ),
    `490` = sum(!is.na(aip_approval_date) & pass_490 == 1, na.rm = T),
    `690` = sum(pass_690 == 1, na.rm = T),
    `07 to 390 %` = as_perc(sum(!is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T) / .N),
    `270 to 690 %` = as_perc(sum(!is.na(aip_approval_date) & pass_490 == 1, na.rm = T) / sum(!is.na(date_of_referral), na.rm = T))
  )]
}

MTD_FUNNEL <- applications %>%
  rollup_dt(c("product", "SKU", "customer_type", "attribution"),
            compute_app_funnel, "Total") %>%
  rename(Product = product, `Customer Type` = customer_type, OIC = attribution) %>%
  arrange(Product, SKU, `Customer Type`, OIC)

colnames(MTD_FUNNEL) [10]  <- "CURRENT_MONTH_LEADS_MOVED_270"

appops_270_MTD <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/appops_270.sql"),
  list(START_DATE = MONTH_START)
))


MTD_FUNNEL <- left_join (MTD_FUNNEL,appops_270_MTD,c('Product','Customer Type'='customer_type','SKU','OIC'))

DAILY_FUNNEL <- applications %>%
  .[as_date(applied_date) == DATE] %>%
  rollup_dt(c("product", "SKU", "customer_type", "attribution"),
            compute_app_funnel, "Total") %>%
  rename(Product = product, `Customer Type` = customer_type, OIC = attribution) %>%
  arrange(Product, SKU, `Customer Type`, OIC)

colnames(DAILY_FUNNEL) [10]  <- "CURRENT_MONTH_LEADS_MOVED_270"

appops_270_DTD <- data.table(read_sql2(
  get_conn(),
  file.path(MIS_PATH, "queries/appops_270.sql"),
  list(START_DATE = DATE)
))

DAILY_FUNNEL <- left_join (DAILY_FUNNEL,appops_270_DTD,c('Product','Customer Type'='customer_type','SKU','OIC'))


MTD_FUNNEL_SUMMARY <- applications %>%
  rollup_dt(c("product", "SKU"), compute_app_funnel, "Total") %>%
  rename(Product = product) %>%
  arrange(Product, SKU)




DAILY_FUNNEL_SUMMARY <- applications %>%
  .[as_date(applied_date) == DATE] %>%
  rollup_dt(c("product", "SKU"), compute_app_funnel, "Total") %>%
  rename(Product = product) %>%
  arrange(Product, SKU)




DAYWISE_SKU_REFERRALS_TREND <- applications %>%
  # .[SKU %in% c("SBI CC", "RBL CC", "Yes Bank CC", "ICICI CC", "HDFC PL", "HDB SBPL",
  #              "Shriram SBPL")] %>%
  .[, app_date := as.character(as_date(applied_date))] %>%
  .[] %>%
  reshape2::dcast(SKU + attribution ~ app_date, margins = T) %>%
  rename(OIC = attribution)

AGENCY_TRACKER_LIST <- lapply(AGENCIES, function(agency) {
  applications %>%
    .[team == agency] %>%
    rollup_dt('SKU', compute_agency_tracker, 'Total') %>%
    arrange(ifelse(SKU == 'Total', 'ZZZ', SKU))
}) %>%
  set_names(AGENCIES)

SCC_REFERRALS_MTD <- applications %>%
  .[ref_oic %in% CIS_OICS$oic & product == 'SCC'] %>%
  .[, .(lead_id, oic = ref_oic, product, date_of_referral)] %>%
  inner_join(CIS_OICS[, .(oic, team)]) %>%
  data.table() %>%
  rollup_dt(
    'team',
    function(dat) dat[, .(Referrals = .N, `S2L %` = as_perc(sum(!is.na(date_of_referral)) / .N))],
    'Total'
    ) %>%
  rename(Team = team)

SCC_REFERRALS_DTD <- applications %>%
  .[as_date(applied_date) == DATE] %>%
  .[ref_oic %in% CIS_OICS$oic & product == 'SCC'] %>%
  .[, .(lead_id, oic = ref_oic, product, date_of_referral)] %>%
  inner_join(CIS_OICS[, .(oic, team)]) %>%
  data.table() %>%
  rollup_dt(
    'team',
    function(dat) dat[, .(Referrals = .N, `S2L %` = as_perc(sum(!is.na(date_of_referral)) / .N))],
    'Total'
    ) %>%
  rename(Team = team)

#' 07 TO 690
DTD_APP_TO_BOOKINGS_FUNNEL <- applications %>%
  .[as_date(applied_date) == DATE] %>%
  rollup_dt(c('product', 'SKU', 'customer_type', 'attribution'), compute_postapp_metrics, 'Total') %>%
  rename(Product = product, `Customer Type` = customer_type, OIC = attribution)

MTD_APP_TO_BOOKINGS_FUNNEL <- applications %>%
  .[is.within.mtd(as_date(applied_date), DATE)] %>%
  rollup_dt(c('product', 'SKU', 'customer_type', 'attribution'), compute_postapp_metrics, 'Total') %>%
  rename(Product = product, `Customer Type` = customer_type, OIC = attribution)

#******************************************************
## 2. Secondary Sale Generated  =======
#******************************************************
loginfo("Referrals MIS - application - generated - Secondary Sale Generated Initated")

chr_secondary <- chr_bys_subscriptions[,c('user_id','payment_date','service_type','oic','team')] %>% 
  rename(
    'applied_date' = 'payment_date',
    'sku' = 'service_type',
    'ref_oic' = 'oic'
  ) %>% 
  mutate(product = sku,
         sku_slug = sku,
         SKU = sku)

applications_ss <- bind_rows(applications,chr_secondary) %>% 
  as.data.table()

applications_ss[order(applied_date), `:=`(
  is_secondary_sale = c(0, rep_len(1, .N - 1)),
  is_primary_sale = c(1, rep_len(0, .N - 1))
),
.(user_id, ref_oic)
]


compute_secondary_sale_metrics <- function(data, value_col) {
  data[, .(
    `Referrals Generated` = sum(!(is_secondary_sale == 0 & product == 'STBL'),
                                na.rm = T),
    `Secondary Sale (SS) Generated` = sum(is_secondary_sale, na.rm = T),
    `SS %` = paste(round(
      sum(is_secondary_sale, na.rm = T) /
        sum(!(is_secondary_sale == 0 & product == 'STBL'), na.rm = T) * 100, 1), "%"),
    `SBPL from SS` = sum(is_secondary_sale == 1 &
                           grepl("shriram", sku_slug, ignore.case = T) &
                           product %in% c("SBPL", "PL"), na.rm = T),
    `ES from SS` = sum(
      is_secondary_sale == 1 &
        grepl("early", sku_slug, ignore.case = T),
      na.rm = T),
    `CashE from SS` = sum(
      is_secondary_sale == 1 &
        grepl("cashe", sku_slug, ignore.case = T),
      na.rm = T),
    `BYS from SS` = sum(
      is_secondary_sale == 1 &
        grepl("bys", sku_slug, ignore.case = T),
      na.rm = T),
    `CHR from SS` = sum(
      is_secondary_sale == 1 &
        grepl("chr", sku_slug, ignore.case = T),
      na.rm = T)
  )] %>%
    gather_("Metric", value_col, names(.))
}

compute_final_ss_table <- function(data) {
  lapply(names(DATE_FILTERS), function(period) {
    data[DATE_FILTERS[[period]](applied_date),
         compute_secondary_sale_metrics(.SD, period)]
  }) %>%
    reduce(full_join, by = "Metric")
}

compute_skuwise_ss_table <- function(data) {
  lapply(names(DATE_FILTERS), function(period) {
    data[DATE_FILTERS[[period]](applied_date) & is_secondary_sale == 1,
         .(count = .N), SKU] %>%
      set_colnames(c("SKU", period))
  }) %>%
    reduce(full_join, by = "SKU")
}

lapply(AGENCIES, function(agency) {
  sec_sale_table <- applications_ss[applications_ss$team == agency,] %>% 
    as.data.table() %>% 
    compute_final_ss_table()

  sec_sale_sku_table <- applications_ss[applications_ss$team == agency,] %>% 
    as.data.table() %>%
    compute_skuwise_ss_table() %>%
    arrange(-MTD)

  final_table_name <- form_agency_table(agency, 'SS', 'summary')
  sku_table_name <- form_agency_table(agency, 'SS', 'skuwise')

  assign(final_table_name, sec_sale_table, .GlobalEnv)
  assign(sku_table_name, sec_sale_sku_table, .GlobalEnv)
})

loginfo("Referrals MIS - application - generated - Secondary Sale Generated Completed")
#******************************************************
## 2.1 210 Reject Reasons =======
#******************************************************
loginfo("Referrals MIS - application - generated - 210 Reject Reasons Initiated")
REJECT_AT_210_SKUS <- applications %>%
  .[product %in% c('HL', 'HLBT', 'LAP', 'SBPL') |
      SKU %in% c('Yes Bank CC', 'INDUSIND PL', 'Shriram PL', 'HDFC PL'), unique(SKU)]

START_COL <- 3

SKUWISE_REJECT_TABLES <- lapply(REJECT_AT_210_SKUS, function(reject_sku) {
  loginfo(str_interp('Running Reject Table for ${reject_sku}'))
  attribution_wise_data <- applications %>%
    .[SKU == reject_sku] %>%
    .[!is.na(not_interested_date) & is.na(date_of_referral)]


  if (nrow(attribution_wise_data) != 0) {
    attribution_wise <- attribution_wise_data %>%
    reshape2::dcast(
      crm_210 ~ attribution,
      value.var = 'lead_id',
      fun.aggregate = n_distinct,
      na.rm = T,
      margins = T
    )
  } else {
    return(create_empty_df(c('Reject Reason', 'System', 'PA', 'Assisted')))
  }


  agency_wise_data <- applications %>%
    .[SKU == reject_sku] %>%
    .[!is.na(not_interested_date) & is.na(date_of_referral) & !is.na(team)]

  if (nrow(agency_wise_data) != 0) {
    agency_wise <- agency_wise_data %>%
      reshape2::dcast(
        crm_210 ~ team,
        value.var = 'lead_id',
        fun.aggregate = n_distinct,
        na.rm = T,
        margins = 'crm_210'
      )
  } else {
      agency_wise <- data.frame(crm_210 = character(0))
  }

  full_join(agency_wise, attribution_wise, by = 'crm_210') %>%
    mutate(crm_210 = as.character(crm_210)) %>%
    mutate(crm_210 = ifelse(is.na(crm_210), 'Others', crm_210)) %>%
    rename(`Reject Reason` = crm_210, Total = `(all)`) %>%
    arrange(ifelse(`Reject Reason` == '(all)', 999999,-Total)) %>%
    select(`Reject Reason`,
           one_of(c('System', 'PA', 'Marketing', 'Assisted')),
           everything(),
           Total)
}) %>%
  set_names(REJECT_AT_210_SKUS)

SKUWISE_REJECT_TABLES %>%
  map2(names(.), function(sku_df, sku) {
    write_excel_table(wb, "210 Reject Reasons", sku_df, 3, START_COL,
                      str_interp('${sku} 210 Reject Reasons'), c('Reject Reason'))

    START_COL <<- START_COL + ncol(sku_df) + 3
  })
loginfo("Referrals MIS - application - generated - 210 Reject Reasons Completed")
#******************************************************
## . S2L and Feedback Received TAT  =======
#******************************************************
loginfo("Referrals MIS - application - generated - S2L and Feedback Received TAT Initated")
compute_s2l_fr_tat_metrics <- function(data) {
  data[, .(
    `Applications` = .N,
    `S2L` = sum(!is.na(date_of_referral), na.rm = T),
    `FR (Feedback Received) Overall` = sum(has_received_feedback == 1, na.rm = T),
    `FR Overall %` = as_perc(
      sum(has_received_feedback == 1, na.rm = T) / sum(!is.na(date_of_referral), na.rm = T)
    ),
    `FR (0-1D)` = sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) < 1,
      na.rm = T),
    `FR (1-4D)` = sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 1 &
        date_diff(feedback_received_date, date_of_referral) < 4,
      na.rm = T),
    `FR (4-7D)` = sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 4 &
        date_diff(feedback_received_date, date_of_referral) < 7,
      na.rm = T),
    `FR (7-10D)` = sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 7 &
        date_diff(feedback_received_date, date_of_referral) < 10,
      na.rm = T),
    `FR (>10D)` = sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 10,
      na.rm = T),
    `FR (0-1D) %` = as_perc(sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) < 1,
      na.rm = T) / sum(has_received_feedback == 1, na.rm = T)),
    `FR (1-4D) %` = as_perc(sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 1 &
        date_diff(feedback_received_date, date_of_referral) < 4,
      na.rm = T) / sum(has_received_feedback == 1, na.rm = T)),
    `FR (4-7D) %` = as_perc(sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 4 &
        date_diff(feedback_received_date, date_of_referral) < 7,
      na.rm = T) / sum(has_received_feedback == 1, na.rm = T)),
    `FR (7-10D) %` =as_perc(sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 7 &
        date_diff(feedback_received_date, date_of_referral) < 10,
      na.rm = T) / sum(has_received_feedback == 1, na.rm = T)),
    `FR (>10D) %` = as_perc(sum(
      has_received_feedback == 1 &
        date_diff(feedback_received_date, date_of_referral) >= 10,
      na.rm = T) / sum(has_received_feedback == 1, na.rm = T)),
    `390 Movement (Overall)` = sum(!is.na(aip_approval_date), na.rm = T),
    `390 Movement (Overall) %` = as_perc(
      sum(!is.na(aip_approval_date), na.rm = T) / sum(!is.na(date_of_referral), na.rm = T)
    ),
    `390 Movement (0-1D)` = sum(
      has_received_feedback == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) < 1,
      na.rm = T),
    `390 Movement (1-4D)` = sum(
      has_received_feedback == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 1 &
        date_diff(aip_approval_date, date_of_referral) < 4,
      na.rm = T),
    `390 Movement (4-7D)` = sum(
      has_received_feedback == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 4 &
        date_diff(aip_approval_date, date_of_referral) < 7,
      na.rm = T),
    `390 Movement (7-10D)` = sum(
      has_received_feedback == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 7 &
        date_diff(aip_approval_date, date_of_referral) < 10,
      na.rm = T),
    `390 Movement (>10D)` = sum(
      has_received_feedback == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 10,
      na.rm = T),
    `390 Movement (0-1D) %` = as_perc(sum(
      is_aip_approved == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) < 1,
      na.rm = T) / sum(is_aip_approved == 1 & !is.na(aip_approval_date), na.rm = T)),
    `390 Movement (1-4D) %` = as_perc(sum(
      is_aip_approved == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 1 &
        date_diff(aip_approval_date, date_of_referral) < 4,
      na.rm = T) / sum(is_aip_approved == 1 & !is.na(aip_approval_date), na.rm = T)),
    `390 Movement (4-7D) %` = as_perc(sum(
      is_aip_approved == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 4 &
        date_diff(aip_approval_date, date_of_referral) < 7,
      na.rm = T) / sum(is_aip_approved == 1 & !is.na(aip_approval_date), na.rm = T)),
    `390 Movement (7-10D) %` =as_perc(sum(
      is_aip_approved == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 7 &
        date_diff(aip_approval_date, date_of_referral) < 10,
      na.rm = T) / sum(is_aip_approved == 1 & !is.na(aip_approval_date), na.rm = T)),
    `390 Movement (>10D) %` = as_perc(sum(
      is_aip_approved == 1 & !is.na(aip_approval_date) &
        date_diff(aip_approval_date, date_of_referral) >= 10,
      na.rm = T) / sum(is_aip_approved == 1 & !is.na(aip_approval_date), na.rm = T))
  )]
}

FR_390_TAT_TABLE_MTD <- applications %>%
  .[DATE_FILTERS[['MTD']](applied_date)] %>%
  rollup_dt(c('SKU', 'attribution'), compute_s2l_fr_tat_metrics, 'Total') %>%
  rename(Attribution = attribution)



SKUWISE_OTP_METRICS_DTD <- otp_logs %>%
  .[as_date(created_at) == DATE] %>%
  when(nrow(.) > 0 ~ .[, .(
    Sent = .N,
    Verified = sum(has_verified_applied),
    `Verified %` = as_perc(mean(has_verified_applied))
  ), .(sku)],
  ~ create_empty_df(c('SKU')))


SKUWISE_OTP_METRICS_MTD <- otp_logs %>%
  when(nrow(.) > 0 ~ .[, .(
    Sent = .N,
    Verified = sum(has_verified_applied),
    `Verified %` = as_perc(mean(has_verified_applied))
  ), .(sku)],
  ~ create_empty_df(c('SKU')))


loginfo("Referrals MIS - application - generated - S2L and Feedback Received TAT Completed")
#*******************************************************************************
#*******************************************************************************
## 2. Write to Excel =======
#*******************************************************************************
#*******************************************************************************
loginfo("Referrals MIS - application - generated - Write to Excel initiated")
APP_SUMM_COL_START <- 3
write_excel_table(wb, "Applications Summary", DAILY_APPLICATIONS,
                  3, APP_SUMM_COL_START, "Daily Referrals", c("Product", "SKU", "Customer Type"))
APP_SUMM_COL_START <- APP_SUMM_COL_START + ncol(DAILY_APPLICATIONS) + 3
write_excel_table(wb, "Applications Summary", MTD_APPLICATIONS,
                  3, APP_SUMM_COL_START, "MTD Referrals", c("Product", "SKU", "Customer Type"))

APP_SUMM_COL_START <- APP_SUMM_COL_START + ncol(MTD_APPLICATIONS) + 3
write_excel_table(wb, "Applications Summary", DAILY_FUNNEL_SUMMARY,
                  3, APP_SUMM_COL_START, "Daily 07 Funnel Summary", c("Product", "SKU"))
APP_SUMM_COL_START <- APP_SUMM_COL_START + ncol(DAILY_FUNNEL_SUMMARY) + 3
write_excel_table(wb, "Applications Summary", MTD_FUNNEL_SUMMARY,
                  3, APP_SUMM_COL_START, "MTD 07 Funnel Summary",
                  c("Product", "SKU"))

APP_SUMM_COL_START <- APP_SUMM_COL_START + ncol(MTD_FUNNEL_SUMMARY) + 3
write_excel_table(wb, "Applications Summary", DAILY_FUNNEL,
                  3, APP_SUMM_COL_START, "Daily 07 Funnel", c("Product", "SKU", "Customer Type", "OIC"))
APP_SUMM_COL_START <- APP_SUMM_COL_START + ncol(DAILY_FUNNEL) + 3
write_excel_table(wb, "Applications Summary", MTD_FUNNEL,
                  3, APP_SUMM_COL_START, "MTD 07 Funnel", c("Product", "SKU", "Customer Type", "OIC"))

write_excel_table(wb, "Agency-wise Applications", DAILY_AGENCYWISE_APPLICATIONS,
                  3, 3, "Daily Agency-wise Referrals", c("SKU", "Customer Type"))
write_excel_table(wb, "Agency-wise Applications", MTD_AGENCYWISE_APPLICATIONS,
                  3, 3 + ncol(DAILY_AGENCYWISE_APPLICATIONS) + 3,
                  "MTD Agency-wise Referrals", c("SKU", "Customer Type"))

write_excel_table(wb, "Day-wise Applications", DAYWISE_SKU_REFERRALS_TREND,
                  5, 5, "Day-wise Applications", c("SKU", "OIC"), c("(all)"))

write_excel_table(wb, "SCC Applications", SCC_REFERRALS_MTD,
                  3, 3, "SCC Applications - MTD", "Team")
write_excel_table(wb, "SCC Applications", SCC_REFERRALS_DTD,
                  3 + nrow(SCC_REFERRALS_DTD) + 3, 3, "SCC Applications - DTD", "Team")

write_excel_table(wb, "Lender Feedback TATs", FR_390_TAT_TABLE_MTD,
                  3, 3, "Lender Feedback TAT MTD", c("SKU", "Attribution"))



map2(AGENCIES, 1:length(AGENCIES), function(agency, i) {
  write_excel_table(wb, "Secondary Sales",
                    get(form_agency_table(agency, 'SS', 'summary')),
                    3, 3 + (i - 1) * 6,
                    agency)

  write_excel_table(wb, "Secondary Sales",
                    get(form_agency_table(agency, 'SS', 'skuwise')),
                    14, 3 + (i - 1) * 6,
                    agency)

})

pmap(list(AGENCY_TRACKER_LIST, AGENCIES, 1:length(AGENCIES)),
     function(agency_df, agency, i) {
  write_excel_table(wb, "Agency-wise Application Tracker", agency_df,
                    3, 3 + (i - 1) * ncol(agency_df) + 3,
                    paste(agency, 'Referrals Tracker'),
                    'SKU'
  )
})

write_excel_table(wb, "AIP Approvals", DAILY_AIP_APPROVALS,
                  3, 3, "AIP Approvals - DTD", c('Product', 'SKU', 'Customer Type'))
write_excel_table(wb, "AIP Approvals", MTD_AIP_APPROVALS,
                  3 + nrow(DAILY_AIP_APPROVALS) + 3, 3, "AIP Approvals - MTD",
                  c('Product', 'SKU', 'Customer Type'))


write_excel_table(wb, "07 to Bookings Funnel", DTD_APP_TO_BOOKINGS_FUNNEL,
                  3, 3, "DTD 07 to Bookings Funnel",
                  c('Product', 'SKU', 'Customer Type'))

write_excel_table(wb, "07 to Bookings Funnel", MTD_APP_TO_BOOKINGS_FUNNEL,
                  3, 3 + ncol(DTD_APP_TO_BOOKINGS_FUNNEL) + 3, "MTD 07 to Bookings Funnel",
                  c('Product', 'SKU', 'Customer Type'))

write_excel_table(wb, "SKU-wise OTP Metrics", SKUWISE_OTP_METRICS_DTD,
                  3, 3, "DTD SKU-wise OTP Metrics",
                  c('sku'))

write_excel_table(wb, "SKU-wise OTP Metrics", SKUWISE_OTP_METRICS_MTD,
                  3, 3 + ncol(SKUWISE_OTP_METRICS_DTD) + 3, "MTD SKU-wise OTP Metrics",
                  c('sku'))



freezePane(wb, "Applications Summary", firstActiveRow = 4)
freezePane(wb, "AIP Approvals", firstActiveRow = 4)
loginfo("Referrals MIS - application - generated - Write to Excel Completed")
loginfo("Referrals MIS - application - generated - Completed")