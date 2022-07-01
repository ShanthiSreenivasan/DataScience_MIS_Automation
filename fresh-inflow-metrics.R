loginfo("Referrals MIS - Fresh Inflow Metrics - Initiated")
#*******************************************************************************
#*******************************************************************************
## 1. M0 Profiles Funnels =======
#*******************************************************************************
#*******************************************************************************

#******************************************************
## 1.1 M0 Yield Metrics  =======
#******************************************************

fresh_inflow_raw[, `:=`(
  yield_stage = case_when(
    progress == "00" | first_call_lead_type == "User" |
      is.na(first_call_lead_type) ~ "00",
    progress %in% c("02") | first_call_stage %in% c("02", "021") |
      is.na(first_call_lead_type) ~ "021",
    progress %in% c("022") | first_call_stage %in% c("021", "022") |
      is.na(first_call_lead_type) ~ "022/3",
    progress %in% c("05") | first_call_stage %in% c("05", "051", "06") |
      is.na(first_call_lead_type) ~ "05",
    progress %in% c("88") | first_call_stage %in% c("88", "89") |
      is.na(first_call_lead_type) ~ "88",
    progress %in% c("07") & is_called == 1 ~ "No Call & 07",
    T ~ "Others"
  )
)]

M0_YIELD <- fresh_inflow_raw %>%
  .[customer_type %in% c("Green", "Amber"), .(
    Inflow = .N,
    Called = sum(is_called, na.rm = T),
    `No Contact (3 Attempts)` = sum(noeff_3attempts, na.rm = T),
    Contacted = sum(is_contacted, na.rm = T),
    Applied = sum(is_applied, na.rm = T),
    `Yield %` = paste(round(
      sum(is_applied == 1 & is_contacted == 1, na.rm = T) /
        sum(is_contacted == 1, na.rm = T) * 100,
      1), "%")
    ),
    .(`Customer Type` = customer_type, `Stage` = yield_stage)] %>%
  .[order(`Customer Type`, Stage)]

M0_DTD_YIELD <- fresh_inflow_raw %>%
  .[profile_date == DATE] %>%
  .[customer_type %in% c("Green", "Amber"), .(
    Inflow = .N,
    Called = sum(is_called, na.rm = T),
    `No Contact (3 Attempts)` = sum(noeff_3attempts, na.rm = T),
    Contacted = sum(is_contacted, na.rm = T),
    Applied = sum(is_applied, na.rm = T),
    `Yield %` = paste(round(
      sum(is_applied == 1 & is_contacted == 1, na.rm = T) /
        sum(is_contacted == 1, na.rm = T) * 100,
      1), "%")
  ),
  .(`Customer Type` = customer_type, `Stage` = yield_stage)] %>%
  .[order(`Customer Type`, Stage)]

#******************************************************
## 1.2 M0 Profiles Progress vs. Called  =======
#******************************************************

compute_progress_funnel <- function(data) {
  data[, .(
    Stock = .N,
    Called = sum(is_called, na.rm = T),
    `Called %` = paste(round(
      sum(is_called, na.rm = T) / .N * 100,
      1), "%"),
    `3+ Attempts - No Contact` = sum(noeff_3attempts, na.rm = T),
    Contacted = sum(is_contacted, na.rm = T),
    `Contacted %` = paste(round(
      sum(is_contacted, na.rm = T) / .N * 100,
      1), "%")
  )]
}

M0_PROGRESS_CALLED_MTD <- fresh_inflow_raw %>%
  .[customer_type %in% c("Green", "Amber")] %>%
  .[is.within.mtd(as_date(profile_date), DATE)] %>%
  rollup_dt(c("customer_type", "progress"), compute_progress_funnel, "Total") %>%
  rename(`Customer Type` = customer_type, `Stage Progressed` = progress)


M0_PROGRESS_CALLED_DTD <- fresh_inflow_raw %>%
  .[customer_type %in% c("Green", "Amber")] %>%
  .[as_date(profile_date) == DATE] %>%
  rollup_dt(c("customer_type", "progress"), compute_progress_funnel, "Total") %>%
  rename(`Customer Type` = customer_type, `Stage Progressed` = progress)


#*******************************************************************************
#*******************************************************************************
## 3. M0 Product Interest Funnels =======
#*******************************************************************************
#*******************************************************************************

prettify_funnel <- function(data) {
  data %>%
    replace_na(list(
      product = "Total", customer_type = "Total", month_orig = "Total"
    )) %>%
    set_colnames(c("Product", "Customer Type","channel","Profile Vintage", "Total",
                   "Pass 021", "Pass 022","Pass 023","Pass 024", "Pass 03",
                   "Pass 03 (System)",  "Pass 05", "Pass 88",
                   "Pass 110", "Pass 07", "Pass 07 (System)", "Pass 270"))
}

PRODUCT_INT_DTD <- prettify_funnel(prod_int_raw_dtd)
PRODUCT_INT_MTD <- prettify_funnel(prod_int_raw_mtd)


#*******************************************************************************
#*******************************************************************************
## 4. M0 Profiles Funnel =======
#*******************************************************************************
#*******************************************************************************

# Core logic in fresh-inflow-profiles.sql

PROFILE_FUNNEL_DTD <- fresh_inflow_funnel_dtd
PROFILE_FUNNEL_MTD <- fresh_inflow_funnel_mtd

#*******************************************************************************
#*******************************************************************************
## 1. Write to Excel =======
#*******************************************************************************
#*******************************************************************************

# Fresh Inflow Funnels
write_excel_table(wb, "Fresh Inflow Metrics", M0_YIELD,
                  3, 3, "M0 Profiles Funnel - MTD")
write_excel_table(wb, "Fresh Inflow Metrics", M0_DTD_YIELD,
                  3, 3 + ncol(M0_YIELD) + 2, "M0 Profiles Funnel - Daily")


# Fresh Inflow - Stage Progressed vs. Calling
write_excel_table(wb, "Fresh Inflow Metrics", M0_PROGRESS_CALLED_DTD,
                  3 + nrow(M0_YIELD) + 3, 3,
                  "M0 Profiles Stage Progresed vs. Called - DTD",
                  c("Customer Type", "Stage Progressed"))
write_excel_table(wb, "Fresh Inflow Metrics", M0_PROGRESS_CALLED_MTD,
                  3 + nrow(M0_YIELD) + 3, 3 + ncol(M0_YIELD) + 2,
                  "M0 Profiles Stage Progresed vs. Called - MTD",
                  c("Customer Type", "Stage Progressed"))

PRD_INT_START_COL <- 3 + ncol(M0_YIELD) + 3 + ncol(M0_DTD_YIELD) + 3

# Product Interest Funnel
write_excel_table(wb, "Fresh Inflow Metrics", PRODUCT_INT_DTD,
                  3, PRD_INT_START_COL,
                  "M0 Product Interest Funnel - DTD",
                  total_cols = c("Product", "Customer Type", "Profile Vintage"),
                  total_values = "Total",
                  withFilter = T)
write_excel_table(wb, "Fresh Inflow Metrics", PRODUCT_INT_MTD,
                  3, PRD_INT_START_COL + ncol(PRODUCT_INT_DTD) + 3,
                  "M0 Product Interest Funnel - MTD",
                  total_cols = c("Product", "Customer Type", "Profile Vintage"),
                  total_values = "Total",
                  withFilter = T)


PROFILE_FUNNEL_START_COL <- PRD_INT_START_COL + ncol(PRODUCT_INT_DTD) + 3 +
  ncol(PRODUCT_INT_MTD) + 3

# Profiles Funnel
write_excel_table(wb, "Fresh Inflow Metrics", PROFILE_FUNNEL_DTD,
                  3, PROFILE_FUNNEL_START_COL,
                  "M0 Profiles Funnel - DTD",
                  total_cols = c("Product", "Customer Type"),
                  total_values = "Total",
                  withFilter = T)
write_excel_table(wb, "Fresh Inflow Metrics", PROFILE_FUNNEL_MTD,
                  3, PROFILE_FUNNEL_START_COL + ncol(PROFILE_FUNNEL_DTD) + 3,
                  "M0 Profiles Funnel - MTD",
                  total_cols = c("Product", "Customer Type"),
                  total_values = "Total",
                  withFilter = T)
loginfo("Referrals MIS - Fresh Inflow Metrics - Completed")
