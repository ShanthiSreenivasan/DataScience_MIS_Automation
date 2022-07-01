loginfo("Referrals MIS - Calling - Metrics - Initiated")
#' Calling Metrics

compute_oic_call_metrics <- function(data) {
  data[, .(
    `Working days`= n_distinct(as_date(log_date),na.rm=T),
    `Total Calls` = .N,
    `Eff` = sum(!dispo %in% c("NC", "CD"), na.rm = T),
    `Eff %` = as_perc(sum(!dispo %in% c("NC", "CD"), na.rm = T) / .N),
    `Average Calls`= (.N/n_distinct(as_date(log_date),na.rm=T)),
    `Average Eff Calls` = (sum(!dispo %in% c("NC", "CD"), na.rm = T)/n_distinct(as_date(log_date),na.rm=T)),
    `NC` = sum(dispo == "NC", na.rm = T),
    `NC %` = as_perc(sum(dispo == "NC", na.rm = T) / .N),
    `CD` = sum(dispo == "CD", na.rm = T),
    `CD %` = as_perc(sum(dispo == "CD", na.rm = T) / .N),
    `NI` = sum(dispo == "NI", na.rm = T),
    `NI %` = as_perc(sum(dispo == "NI", na.rm = T) / .N),
    `CB` = sum(dispo == "CB", na.rm = T),
    `CB %` = as_perc(sum(dispo == "CB", na.rm = T) / .N),
    `NE` = sum(dispo == "NE", na.rm = T),
    `NE %` = as_perc(sum(dispo == "NE", na.rm = T) / .N),
    `PTP` = sum(dispo == "PTP", na.rm = T),
    `DNC` = sum(dispo == "DNC", na.rm = T),
    `110 %` = as_perc(sum(stage %in% c("110", "210")) / .N),
    `CHR PTP` = sum(grepl('chr ptp', crm_status_text, ignore.case = T)),

    `Missed CHR PTP`=n_distinct(lead_id[missed_ptp_chr == 1], na.rm = T),

    `BYS PTP` = sum(grepl('bys ptp', crm_status_text, ignore.case = T)),

    `Missed BYS PTP`=n_distinct(lead_id[missed_ptp_bys == 1], na.rm = T),

    `HL Calls` = sum(product == 'HL', na.rm = T)
  )]
}

compute_form_submits <- function (data) {
  data [, .(form_submits = .N,
            CC_form_submits = sum(grepl('CC', product, ignore.case = T)),
            PL_form_submits = sum(grepl('PL', product, ignore.case = T)),
            Others_form_submits = sum(!product %in% c('CC', 'PL'))
            )]
}


compute_app_metrics <- function(data) {
  data[, .(
    `Referrals` = .N,
    `CC Referrals`=sum(product %in% c('CC')),
    `PL Referrals`=sum(product %in% c('PL') & !SKU %in% c('Shriram PL')),
    `STBL Referrals`=sum(product %in% c('STBL')),
    `SBPL Referrals`=sum(product %in% c('SBPL') | SKU %in% c('Shriram PL')),
    `LOC Referrals`=sum(product %in% c('LOC')),
    `BL Referrals`=sum(product %in% c('BL')),
    `Secondary Sale` = sum(is_secondary_sale == 1, na.rm = T),
    `SS %` = as_perc(sum(is_secondary_sale == 1, na.rm = T) / .N),
    `S2L %` = as_perc(sum(!is.na(date_of_referral)) / .N),
    `NI (210) %` = as_perc(sum(!is.na(not_interested_date)) / .N),
    `240 %` = as_perc(sum(appops_status_code == '240') / .N),
    `270` = sum(appops_status_code %in% c('270')),
    `AIP` = sum(!is.na(aip_approval_date), na.rm = T),
    `AIP (API SKUs)` = sum(!is.na(aip_approval_date) & SKU %in% c('SBI CC', 'RBL CC', 'RBL PL'), na.rm = T),
    `AIP (Non-API SKUs)` = sum(!is.na(aip_approval_date) & !SKU %in% c('SBI CC', 'RBL CC', 'RBL PL'), na.rm = T),
    `AIP %` = as_perc(sum(!is.na(aip_approval_date), na.rm = T) / .N),
    `AIP (API SKUs) %` = as_perc(sum(!is.na(aip_approval_date) & SKU %in% c('SBI CC', 'RBL CC', 'RBL PL'), na.rm = T) / .N),
    `AIP (Non-API SKUs) %` = as_perc(sum(!is.na(aip_approval_date) & !SKU %in% c('SBI CC', 'RBL CC', 'RBL PL'), na.rm = T) / .N),
    `Docs Completed` = sum(document_status %in% c('completed', 'partial')),
    `350/60/70` = sum(appops_status_code %in% c('350', '360', '370')),
    `380` = sum(appops_status_code %in% c('380')),
    `390` = sum(!is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T),
    `390 %` = as_perc(
      sum(!is.na(aip_approval_date) & is_aip_approved == 1, na.rm = T) /
        .N
    ),
    #Change Begin
    `399` = sum(appops_status_code %in% c('399')),
    `400` = sum(appops_status_code %in% c('400')),
    #Change End
    `450/60/70` = sum(appops_status_code %in% c('450', '460', '470')),
    `points` = sum(points_sku),
    `480` = sum(appops_status_code %in% c('480')),
    `490` = sum(!is.na(aip_approval_date) & pass_490 == 1, na.rm = T),
    `490 %` = as_perc(
      sum(!is.na(aip_approval_date) & pass_490 == 1, na.rm = T) / .N)

  )]
}

compute_chr_bys_sms_metrics <- function(data) {
  data[, .(
    `CHR Sms Sent` = sum(service_type %in% c('CHR', 'CHRA'), na.rm = T),
    `BYS Sms Sent` = sum(service_type == 'BYS', na.rm = T)
  )]
}

compute_chr_bys_sub_metrics <- function(data) {
  data[, .(
    `CHR Subscriptions` = sum(service_type %in% c('CHR', 'CHRA'), na.rm = T),
    `BYS Subscriptions` = sum(service_type == 'BYS', na.rm = T)
  )]
}

compute_chr_bys_ptp_metrics <- function(data) {
  data[, .(
    `CHR PTP` = sum(grepl('chr ptp', crm_status_text, ignore.case = T)),
    `BYS PTP` = sum(grepl('bys ptp', crm_status_text, ignore.case = T))
  )]
}

compute_otp_metrics <- function(data) {
  data[, .(
    `OTP Sent` = .N,
    `OTP Verified` = sum(has_verified_applied, na.rm = T)
  )]
}


loginfo("Referrals MIS - Calling - Metrics - OIC-wise metrics - Initiated")
#' OIC-wise metrics
lapply(AGENCIES, function(agency) {
  lapply(names(DATE_FILTERS), function(period) {
    calls_base <- calls_raw[team == agency & DATE_FILTERS[[period]](log_date)]
    apps_base <- applications[team == agency & DATE_FILTERS[[period]](applied_date)]
    chr_bys_sms_base <- chr_bys_sms[team == agency & DATE_FILTERS[[period]](updated_at)]
    chr_bys_sub_base <- chr_bys_subscriptions[team == agency & DATE_FILTERS[[period]](payment_date)]
    otp_logs <- otp_logs[team == agency & DATE_FILTERS[[period]](created_at)]
    form_submits <- oic_form_submitts_mtd[team == agency & DATE_FILTERS[[period]](submitted_at)]

    #' Calling Metrics
    if (nrow(calls_base) == 0) {
      calls_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      calls_summary <- calls_base[, compute_oic_call_metrics(.SD),
                                  .(name = name, oic = oic,oic_type=oic_type)]
    }

    #' Application Metrics
    if (nrow(apps_base) == 0) {
      apps_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      apps_summary <- apps_base[, compute_app_metrics(.SD),
                                .(name = name, oic = ref_oic,oic_type=oic_type)]
    }

    #' CHR/BYS Sms Sent Metrics
    if (nrow(chr_bys_sms_base) == 0) {
      chr_bys_sms_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      chr_bys_sms_summary <- chr_bys_sms_base[, compute_chr_bys_sms_metrics(.SD),
                                .(name = name, oic = oic,oic_type=oic_type)]
    }

    #' CHR/BYS Subscription Metrics
    if (nrow(chr_bys_sub_base) == 0) {
      chr_bys_sub_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      chr_bys_sub_summary <- chr_bys_sub_base[, compute_chr_bys_sub_metrics(.SD),
                                .(name = name, oic = oic,oic_type=oic_type)]
    }

    #' OTP Metrics
    if (nrow(otp_logs) == 0) {
      otp_metrics_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      otp_metrics_summary <- otp_logs[, compute_otp_metrics(.SD),
                                .(name = name, oic = oic,oic_type=oic_type)]
    }

    #' form submits Metrics
    if (nrow(form_submits) == 0) {
      form_submits_summary <- data.frame(oic = character(0), name = character(0),oic_type=character(0))
    } else {
      form_submits_summary <- form_submits[, compute_form_submits(.SD),
                                      .(name = name, oic = oic,oic_type=oic_type)]
    }

    final_table <- left_join(calls_summary, apps_summary, by = c("name", "oic","oic_type")) %>%
      left_join(form_submits_summary, by = c("name", "oic","oic_type")) %>%
      left_join(chr_bys_sms_summary, by = c("name", "oic","oic_type")) %>%
      left_join(chr_bys_sub_summary, by = c("name", "oic","oic_type")) %>%
      left_join(otp_metrics_summary, by = c("name", "oic","oic_type")) %>%
      rename(Name = name, OIC = oic,oic_type=oic_type)


    final_table_name <- form_agency_table(agency, period, "OIC_CALLS")

    assign(final_table_name, final_table, .GlobalEnv)
  })
})
loginfo("Referrals MIS - Calling - Metrics - OIC-wise metrics - Completed")

lapply(names(DATE_FILTERS), function(period) {
  calls_base <-
    calls_raw[team %in% AGENCIES & DATE_FILTERS[[period]](log_date)]
  apps_base <-
    applications[team %in% AGENCIES & DATE_FILTERS[[period]](applied_date)]
  chr_bys_sms_base <-
    chr_bys_sms[team %in% AGENCIES & DATE_FILTERS[[period]](updated_at)]
  chr_bys_sub_base <-
    chr_bys_subscriptions[team %in% AGENCIES & DATE_FILTERS[[period]](payment_date)]
  otp_logs_base <-
    otp_logs[team %in% AGENCIES & DATE_FILTERS[[period]](created_at)]
  form_submits_base <-
    oic_form_submitts_mtd[team %in% AGENCIES & DATE_FILTERS[[period]](submitted_at)]



  if (nrow(calls_base) == 0) {
    calls_summary <- data.frame(team = character(0))
  } else {
    calls_summary <- rollup_dt(calls_base, "team", compute_oic_call_metrics, "Total")
  }

  if (nrow(apps_base) == 0) {
    apps_summary <- data.frame(team = character(0))
  } else {
    apps_summary <- rollup_dt(apps_base, "team", compute_app_metrics, "Total")
  }

  #' CHR/BYS Sms Sent Metrics
  if (nrow(chr_bys_sms_base) == 0) {
    chr_bys_sms_summary <- data.frame(team = character(0))
  } else {
    chr_bys_sms_summary <- rollup_dt(
      chr_bys_sms_base, 'team', compute_chr_bys_sms_metrics, 'Total'
    )
  }

  #' CHR/BYS Subscription Metrics
  if (nrow(chr_bys_sub_base) == 0) {
    chr_bys_sub_summary <- data.frame(team = character(0))
  } else {
    chr_bys_sub_summary <- rollup_dt(
      chr_bys_sub_base, 'team', compute_chr_bys_sub_metrics, 'Total'
    )
  }

  #' OTP Metrics
  if (nrow(otp_logs_base) == 0) {
    otp_metrics_summary <- data.frame(team = character(0))
  } else {
    otp_metrics_summary <- rollup_dt(
      otp_logs_base, 'team', compute_otp_metrics, 'Total'
    )
  }

# form_submitts
  if (nrow(form_submits_base) == 0) {
    form_submits_base_summary <- data.frame(team = character(0))
  } else {
    form_submits_base_summary <- rollup_dt(
      form_submits_base, 'team', compute_form_submits, 'Total'
    )
  }


    final_table <- left_join(calls_summary, apps_summary, by = c("team")) %>%
      left_join(form_submits_base_summary, by = c("team")) %>%
      left_join(chr_bys_sms_summary, by = c("team")) %>%
      left_join(chr_bys_sub_summary, by = c("team")) %>%
      left_join(otp_metrics_summary, by = c("team")) %>%
    rename(Team = team)

  final_table_name <-
    paste("OVERALL", toupper(period), "CALLS", sep = "_")

  assign(final_table_name, final_table, .GlobalEnv)
})


generate_excel_table_for_agency <- function(agency) {
  daily_calls_table <- get(form_agency_table(agency, 'Daily', 'OIC_CALLS'))
  mtd_calls_table <- get(form_agency_table(agency, 'MTD', 'OIC_CALLS'))

  write_excel_table(wb, agency, daily_calls_table,
                    3, 3, "OIC-wise Daily Calls")
  write_excel_table(wb, agency, mtd_calls_table,
                    3, 3 + ncol(daily_calls_table) + 3, "OIC-wise MTD Calls")
  loginfo(str_interp("Referrals MIS - ${agency} Calling Metrics Exported to Excel"))
}

#*******************************************************************************
#*******************************************************************************
## 3. Write to Excel =======
#*******************************************************************************
#*******************************************************************************

write_excel_table(wb, "Calling Summary", OVERALL_DAILY_CALLS,
                  3, 3, "Agency-wise Daily Calls", "Team")
write_excel_table(wb, "Calling Summary", OVERALL_MTD_CALLS,
                  3 + nrow(OVERALL_DAILY_CALLS) + 3, 3, "Agency-wise MTD Calls", "Team")

lapply(AGENCIES, generate_excel_table_for_agency)

loginfo("Referrals MIS - Calling - Metrics - Completed")

