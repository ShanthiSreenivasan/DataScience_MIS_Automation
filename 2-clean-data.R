 
#*******************************************************************************
#*******************************************************************************
## 1. Clean Applications Data =======
#*******************************************************************************
#*******************************************************************************

loginfo("Referrals MIS - 2-clean-data - Initiated")
'%nin%' <- Negate(`%in%`)  #rajan changes 2/19/2021
applications_raw[, `:=`(
  ref_oic = case_when(
    #Rajan changes start
    utm_term %in% c("single-click") ~ "PORTFOLIO",
      # grepl("(stbl-cashe|stbl-moneyview|STBL-EARLYSALARYNINETY|STBL-EARLYSALARYNINETY_androidApp)", utm_source, ignore.case = T) ~ "PORTFOLIO",
      applied_oic %in% c("SingleClick") ~ "PORTFOLIO",
      #Rajan changes end
    applied_oic %in% c("lead-reg") ~ "Regeneration",
    !applied_oic %in% c("System", "PORTFOLIO", "andriodApp", "referralCustomerEngagement") ~ applied_oic,
    applied_oic == "System" &
      !grepl("(monitoring|^email-[RAG]2[RAG][di]?|^email-chra|^email-cis-(310|350|4050))", utm_source, ignore.case = T) &
      grepl("(email|sms)", utm_source, ignore.case = T) ~ "PORTFOLIO",
    #Nirmal-Addition
    applied_oic == "andriodApp" &
      grepl("(PA|pa)", utm_source, ignore.case = T) ~ "PORTFOLIO",
    #Nirma-Cheanges End
    applied_oic == "System" &
      (utm_source %in% PRELOGIN_UTMS$utm |
         grepl("(widget$|prelogin)", utm_content, ignore.case = T)) ~ "Marketing",
    T ~ "System"
  ),
  #rajan changes start 2/19/2021
  #attribution = case_when(applied_oic %in% c("lead-reg") ~ "Regeneration",
  attribution = case_when(
    #Rajan changes start
    utm_term %in% c("single-click") ~ "PA",
      # grepl("(stbl-cashe|stbl-moneyview|STBL-EARLYSALARYNINETY|STBL-EARLYSALARYNINETY_androidApp)", utm_source, ignore.case = T) ~ "PA",
    applied_oic %in% c("SingleClick") ~ "PA",
    #Rajan changes end
  (applied_oic %in% c("lead-reg") & first_oic %nin% c("PORTFOLIO","andriodApp","System"))~ "Regeneration_Others",
  (applied_oic %in% c("lead-reg") & first_oic %in% c("andriodApp","System"))~ "Regeneration_STP",
  (applied_oic %in% c("lead-reg") & first_oic %in% c("PORTFOLIO"))~ "Regeneration_PA",
  #rajan changes end 2/19/2021
    !applied_oic %in% c("System", "PORTFOLIO", "andriodApp", "referralCustomerEngagement") ~ "Assisted",
    applied_oic == "System" &
      !grepl("(monitoring|^email-[RAG]2[RAG][di]?|^email-chra|^email-cis-(310|350|4050))", utm_source, ignore.case = T) &
      grepl("(email|sms)", utm_source, ignore.case = T) ~ "PA",
    #Nirmal changes start
    applied_oic == "andriodApp" &
      grepl("(PA|pa)", utm_source, ignore.case = T) ~ "PA",
    #Nirmal changes end
    applied_oic == "System" &
      (utm_source %in% PRELOGIN_UTMS$utm |
         grepl("(widget$|prelogin)", utm_content, ignore.case = T)) ~ "Marketing",
    T ~ "System"
  ) %>% factor(c("System", "Assisted","Regeneration_Others","Regeneration_STP","Regeneration_PA", "PA", "Marketing", "Total"), ordered = T),
  SKU = sku
)][, `:=`(
  attr_class = factor(ifelse(
    attribution %in% c("Assisted", "PA","Regeneration_Others","Regeneration_STP","Regeneration_PA","System"), #rajan changes 2/19/2021
    "Non-Marketing",
    "Marketing"), c("Non-Marketing", "Marketing"), ordered = T)
)]

aip_approval_raw[, `:=`(
  ref_oic = case_when(
    #Rajan changes start
    utm_term %in% c("single-click") ~ "PORTFOLIO",
    # grepl("(stbl-cashe|stbl-moneyview|STBL-EARLYSALARYNINETY|STBL-EARLYSALARYNINETY_androidApp)", utm_source, ignore.case = T) ~ "PORTFOLIO",
    applied_oic %in% c("SingleClick") ~ "PORTFOLIO",
    #Rajan changes end
    applied_oic %in% c("lead-reg") ~ "Regeneration",
    !applied_oic %in% c("System", "PORTFOLIO", "andriodApp", "referralCustomerEngagement") ~ applied_oic,
    applied_oic == "System" &
      !grepl("(monitoring|^email-[RAG]2[RAG][di]?|^email-chra|^email-cis-(310|350|4050))", utm_source, ignore.case = T) &
      grepl("(email|sms)", utm_source, ignore.case = T) ~ "PORTFOLIO",
    #Nirmal-Addition
    applied_oic == "andriodApp" &
      grepl("(PA|pa)", utm_source, ignore.case = T) ~ "PORTFOLIO",
    #Nirma-Cheanges End
    applied_oic == "System" &
      (
        utm_source %in% PRELOGIN_UTMS$utm |
          grepl("(widget$|prelogin)", utm_content, ignore.case = T)
      ) ~ "Marketing",
    T ~ "System"
  ),
  attribution = case_when(
    #Rajan changes start
    utm_term %in% c("single-click") ~ "PA",
    # grepl("(stbl-cashe|stbl-moneyview|STBL-EARLYSALARYNINETY|STBL-EARLYSALARYNINETY_androidApp)", utm_source, ignore.case = T) ~ "PA",
    applied_oic %in% c("SingleClick") ~ "PA",
    #Rajan changes end
    applied_oic %in% c("lead-reg") ~ "Regeneration",
    !applied_oic %in% c("System", "PORTFOLIO", "andriodApp","referralCustomerEngagement") ~ "Assisted",
    applied_oic == "System" &
      !grepl("(monitoring|^email-[RAG]2[RAG][di]?|^email-chra|^email-cis-(310|350|4050))", utm_source, ignore.case = T) &
      grepl("(email|sms)", utm_source, ignore.case = T) ~ "PA",
    #Nirmal-Addition
    applied_oic == "andriodApp" &
      grepl("(PA|pa)", utm_source, ignore.case = T) ~ "PA",
    #Nirma-Cheanges End
    applied_oic == "System" &
      (
        utm_source %in% PRELOGIN_UTMS$utm |
          grepl("(widget$|prelogin)", utm_content, ignore.case = T)
      ) ~ "Marketing",
    T ~ "System"
  ) %>% factor(c(
    "System", "Assisted", "PA","Regeneration","Marketing", "Total"
  ), ordered = T),
  SKU = paste(lender, product)
)]


aip_approved_cases <- aip_approval_raw %>%
  left_join(REFERRALS_OIC, by = c("ref_oic" = "oic")) %>%
  data.table()

applications <- applications_raw %>%
  .[is.within.mtd(as_date(applied_date), DATE)] %>%
  left_join(REFERRALS_OIC, by = c("ref_oic" = "oic")) %>%
  data.table()

applications[order(applied_date), `:=`(
  is_secondary_sale = c(0, rep_len(1, .N - 1)),
  is_primary_sale = c(1, rep_len(0, .N - 1))
  ),
  .(user_id, ref_oic)
]

applications[, `:=`(
  is_aip_approved = case_when(
    !is.na(aip_approval_date) & !SKU %in% c('Early Salary STBL', 'RBL CC', 'SBI CC', 'ICICI CC') ~ 1,
    SKU == 'Early Salary STBL' & pass_390 == 1 ~ 1,
    SKU == 'CashE STBL' & pass_390 == 1 ~ 1,
    SKU == 'RBL CC' & is_rbl_cc_api_approved == 1 ~ 1,
    SKU == 'SBI CC' & is_sbi_cc_api_approved == 1 ~ 1,
    SKU == 'ICICI CC' & pass_490 == 1 ~ 1,
    T ~ 0
  )
)]

applications_1<-applications
#*******************************************************************************
#*******************************************************************************
## 2. Fresh Inflow Data =======
#*******************************************************************************
#*******************************************************************************

fresh_inflow_raw[, `:=`(
  noeff_3attempts = as.integer(attempts >= 3 & eff_attempts == 0)
)]


#*******************************************************************************
#*******************************************************************************
## 3. Calls Data =======
#*******************************************************************************
#*******************************************************************************
 
calls_raw <- REFERRALS_OIC[, .(oic, name, team,oic_type)][calls_raw, on="oic"]

loginfo("Referrals MIS - 2-clean-data - Preprocessing Completed")


#*******************************************************************************
#*******************************************************************************
## 4. Export to Excel =======
#*******************************************************************************
#*******************************************************************************

write_excel_table(wb, "OIC App Downloads", APP_DOWNLOADS_DTD,
                  3, 3, "App Downloads - Daily Summary", c("OIC"))
write_excel_table(wb, "OIC App Downloads", APP_DOWNLOADS_MTD,
                  3, 3 + ncol(APP_DOWNLOADS_DTD) + 3,
                  "App Downloads - MTD Summary", c("OIC"))

write_excel_table(wb, "OIC Logins", OIC_LOGIN_DATA,
                  3, 3,
                  "OIC Logins - MTD Summary")

loginfo("Referrals MIS - 2-clean-data - Written tables to Excel")

#*******************************************************************************
#*******************************************************************************
## 5. Write & Upload Dumps to S3 =======
#*******************************************************************************
#*******************************************************************************

applications_filename <- file.path(DATA_PATH, str_interp("Referrals Dump - ${DATE}.csv"))

applications %>%
  select(-applied_oic, -utm_content, -SKU,
         -attr_class) %>%
  fwrite(applications_filename, dateTimeAs = "write.csv")

fresh_inflow_filename <-file.path(DATA_PATH, str_interp("M0 Profiles Inflow - ${DATE}.csv"))
fwrite(fresh_inflow_raw, fresh_inflow_filename, dateTimeAs = "write.csv")
aws.s3::put_object(
  file = fresh_inflow_filename,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(fresh_inflow_filename))
)


zipfile <- file.path(DATA_PATH, str_interp("07 Dump - ${DATE}.zip"))
zip(zipfile = zipfile, files = c(applications_filename))
aws.s3::put_object(
  file = zipfile,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(zipfile))
)
aws.s3::put_object(
  file = zipfile,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(zipfile))
)
aws.s3::put_object(
  file = zipfile,
  bucket = MARKETING_REPORTS_BUCKET,
  object = file.path(MARKETING_REPORTS_KEY, basename(zipfile))
)

loginfo("Referrals MIS - 2-clean-data - Exported dumps to S3")

#*******************************************************************************
## 5.1 CHR/BYS dumps =======
#*******************************************************************************

chr_bys_sms_filename <- file.path(DATA_PATH, str_interp('CHR-BYS SMS Sent Details - ${DATE}.csv'))
fwrite(chr_bys_sms, chr_bys_sms_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = chr_bys_sms_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(chr_bys_sms_filename))
)


chr_bys_subs_filename <- file.path(DATA_PATH, str_interp('CHR-BYS Subscription Details - ${DATE}.csv'))
fwrite(chr_bys_subscriptions_overall, chr_bys_subs_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(chr_bys_subs_filename))
)
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = CIS_REPORTS_BUCKET,
  object = file.path(CIS_REPORTS_KEY, basename(chr_bys_subs_filename))
)
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = MARKETING_REPORTS_BUCKET,
  object = file.path(MARKETING_REPORTS_KEY, basename(chr_bys_subs_filename))
)
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(chr_bys_subs_filename))
)
################# Nirmal Addition
idpp_subs_filename <- file.path(DATA_PATH, str_interp('IDPP Subscription Details - ${DATE}.csv'))
fwrite(idpp_subscriptions_overall, idpp_subs_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = idpp_subs_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(idpp_subs_filename))
)
aws.s3::put_object(
  file = idpp_subs_filename,
  bucket = CIS_REPORTS_BUCKET,
  object = file.path(CIS_REPORTS_KEY, basename(idpp_subs_filename))
)
aws.s3::put_object(
  file = idpp_subs_filename,
  bucket = MARKETING_REPORTS_BUCKET,
  object = file.path(MARKETING_REPORTS_KEY, basename(idpp_subs_filename))
)
aws.s3::put_object(
  file = idpp_subs_filename,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(idpp_subs_filename))
)


chr_bys_subs_filename <- file.path(DATA_PATH, str_interp('BHRL Subscription Details - ${DATE}.csv'))
fwrite(bhrl_subscriptions_overall, chr_bys_subs_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(chr_bys_subs_filename))
)
aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = CIS_REPORTS_BUCKET,
  object = file.path(CIS_REPORTS_KEY, basename(chr_bys_subs_filename))
)

aws.s3::put_object(
  file = chr_bys_subs_filename,
  bucket = REPORTS_BUCKET,
  object = file.path(REPORTS_KEY, basename(chr_bys_subs_filename))
)


#############################
chr_bys_interest_filename <- file.path(DATA_PATH, str_interp('CHR-BYS Interested Customers Dump - ${DATE}.csv'))
fwrite(chr_bys_interested_base, chr_bys_interest_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = chr_bys_interest_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(chr_bys_interest_filename))
)


chr_bys_070_filename <- file.path(DATA_PATH, str_interp('CHR-BYS 070 Dump - ${DATE}.csv'))
fwrite(chr_bys_070_base, chr_bys_070_filename, dateTimeAs = 'write.csv')
aws.s3::put_object(
  file = chr_bys_070_filename,
  bucket = ASSISTED_REPORTS_BUCKET,
  object = file.path(ASSISTED_REPORTS_KEY, basename(chr_bys_070_filename))
)

loginfo("Referrals MIS - Clean data - Completed")
