
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

#options(digits = 0)
REF_TARGET <- fread('C:\\R\\Referrals_monthly_target_Feb.csv')%>%
  melt(id.vars = c('customer_type',"product", 'SKU'),
       measure.vars = c("System",'PA',"Allset","Regeneration","Assisted")) %>%
  dplyr::rename('Attribution' = variable,
         'Monthly_Plan' = value)
REF_MIS <- fread('C:\\R\\Referrals Dump - 2022-02-23.csv') %>%
  melt(id.vars = c('customer_type',"product", 'sku')),
       measure.vars = c("System",'PA',"Allset","Regeneration_STP","Assisted") %>%
  dplyr::rename('Attribution' = variable,
                'Monthly_Plan' = value)
jointdataset <-merge (REF_MIS, REF_TARGET, by= 'sku', 'customer_type') %>% mutate('DTD Plan' = round(`Monthly_Plan` / 30) %>%
  mutate_if(is.character, toupper))

class(REF_TARGET$Monthly_Plan)
#assisted_date_past <- data.table(read_sql2(
 # get_conn(),
  #file.path(MIS_PATH, "queries/assisted_working_days.sql"),
  #list(START_DATE = MONTH_START, END_DATE = DATE)
#))

#Always Assited should be in last
attributions <- c("System",'PA',"Allset","Regeneration","Assisted")
# attributions <- c("Allset")
days.month <- as.numeric(days_in_month)
days.past <- as.numeric(DAYS_PAST)

COL_START <- 3
# a=list()

for(team in attributions){



  mtd_target <- REF_TARGET[REF_TARGET$attribution == team,]
  mtd_target <- mtd_target[mtd_target$`Monthly Plan` != 0,]

  applications <- as.data.frame(applications)

  days.month <- case_when(team == 'Assisted' ~ 25, T ~ as.numeric(DAYS_IN_MONTH))
  days.past <- case_when(team == 'Assisted' ~ assisted_date_past[[1]], T ~ as.numeric(DAYS_PAST))
  days.left <- days.month - days.past

  days.month <- case_when(team == 'Allset' ~ 25, T ~ as.numeric(DAYS_IN_MONTH))
  days.past <- case_when(team == 'Allset' ~ assisted_date_past[[1]], T ~ as.numeric(DAYS_PAST))
  days.left <- days.month - days.past


  if(team %in% c("Allset")){

    app_system <- applications[applications$team %in% team,]


  }
  #rajan changes start 09/08/2021
  if(team %in% c("Regeneration")){

    app_system <- applications[applications$attribution %in% c('Regeneration_Others','Regeneration_PA','Regeneration_STP'),]

    #rajan changes end 09/08/2021
  }else{

    app_system <- applications[applications$attribution == team,]

  }


  #To make all varible as lower
  app_system$customer_type <- toupper(app_system$customer_type)
  app_system$product <- toupper(app_system$product)
  app_system$SKU <- toupper(app_system$SKU)
  # a[[team]]]<-app_system%>%filter (as_date(applied_date) == DATE)

  app_system <- app_system %>%
    as.data.table() %>%
    .[,.(
      # 'DRR' = (REF_TARGET$`Monthly Plan`-.N)/as.numeric(MONTH_END)-days.left, #rajan changes 09/08/2021
      'DTD Actual' = n_distinct(lead_id[as_date(app_system$applied_date) == DATE],na.rm = T),
      'MTD Actual' = .N,
      'MTD Estimate' = round( (.N / days.past) * days.month),0),
      list(customer_type,product, SKU)] %>%
    .[order(customer_type,product, SKU)] %>%
    left_join(y = mtd_target, by = c('customer_type','product','SKU')) %>%
    mutate(
      'MTD Plan' = round((`Monthly Plan`/ days.month) * days.past,0),
      'DTD %' = as_perc(`DTD Actual` / `DTD Plan`),
      'MTD %' = as_perc(`MTD Actual` / `MTD Plan`)
      # 'Backlog' = (round((`Monthly Plan`/ days.month) * days.past,0)) -.N #rajan changes 09/08/2021
    ) %>%
    rename(
      'Customer Type' = customer_type,
      'Product' = product) %>%
    select(
      `Customer Type`,Product, SKU,`DTD Actual`, `DTD Plan`,  `DTD %`,
      `MTD Actual`, `MTD Plan`,`MTD %`,`MTD Estimate`, `Monthly Plan`) %>%
    as.data.table()

  #rajan changes start 09/08/2021
  if(nrow(app_system)!=0){
    Backlog = as.numeric(app_system$`MTD Plan`- app_system$`MTD Actual`)
    # 'Backlog' = (round((`Monthly Plan`/ days.month) * days.past,0)) -.N,
    DRR = as.numeric(round(((app_system$`Monthly Plan`-app_system$`MTD Actual`)/as.numeric(days.month-days.left)),0))

    b1<-as.data.frame(cbind(Backlog,DRR))
    app_system<-as.data.frame(cbind(app_system,b1))
  }
  #rajan changes end 09/08/2021

  sku_list = REF_TARGET[REF_TARGET$attribution == team & REF_TARGET$`Monthly Plan` != 0,]
  skus  = paste(sku_list$customer_type, sku_list$product, sku_list$SKU)
  app_system$sku_list <- paste(app_system$`Customer Type`, app_system$Product, app_system$SKU)

  app_system <- app_system[app_system$sku_list %in% skus,]

  app_system$sku_list <- NULL


  print(app_system %>% dim())
  write_excel_table(wb, "C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\Referrals Forecast.csv", app_system, 3, COL_START, team, c("Product", "SKU", "Customer Type"))
  COL_START = COL_START + ncol(app_system) + 3
  column_no<-COL_START
}

#rajan changes start 09/08/2021
#S2L MTD funnel
compute_app_funnel_S2L <- function(data) {
  data[, .(
    `Applications` = .N,
    `Moved to 270` = sum(!is.na(date_of_referral), na.rm = T),
    `Moved to 390` = sum(!is.na(aip_approval_date) | is_aip_approved == 1, na.rm = T)
    # `Moved_270` = sum(!is.na(date_of_referral), na.rm = T)
  )]
}

MTD_FUNNEL_S2L <- applications_1 %>%
  rollup_dt(c("product", "SKU", "attribution"),
            compute_app_funnel_S2L, "Total") %>%
  rename(Product = product, OIC = attribution) %>%
  arrange(Product, SKU)

if(nrow(MTD_FUNNEL_S2L)!=0 ){

  # names(MTD_FUNNEL_S2L)[which(names(MTD_FUNNEL_S2L)[]=='product')]='Product'
  # names(MTD_FUNNEL_S2L)[which(names(MTD_FUNNEL_S2L)[]=='attribution')]='OIC'

  g<-MTD_FUNNEL_S2L[which(MTD_FUNNEL_S2L[,3]=='Total'),]
  MTD_FUNNEL_S2L_FINAL<-g[-which(g[,2]=='Total'),]


  if(nrow(appops_270_MTD)!=0){
    if(names(appops_270_MTD)[which(names(appops_270_MTD)[]=='customer_type')]=='customer_type'){
      appops_270_MTD[[which(names(appops_270_MTD)[]=='customer_type')]]<-NULL
    }

    appops_270_MTD<-unique(appops_270_MTD)

    appops_270_MTD[,3]<-NULL

    f1<-appops_270_MTD%>%
      group_by(Product,SKU)%>%
      summarise(sum(MOVED_270))

    appops_270_MTD_1<-f1[-which(f1[,2]=='Total'),]

    names(appops_270_MTD_1)[which(names(appops_270_MTD_1)=='sum(MOVED_270)')]='MOVED_270'
  }else{
    appops_270_MTD_1<-data.frame('Product'=character(0),'SKU'=character(0),'MOVED_270'=numeric(0))
  }

  MTD_FUNNEL_S2L_2 <- left_join (MTD_FUNNEL_S2L_FINAL,appops_270_MTD_1,c('Product','SKU'))

  SKU_LENDER_LIST<-unique(applications_1%>%
                            select(lender,sku))

  if(nrow(SKU_LENDER_LIST)!=0){
    SKU_270<-SKU_LENDER_LIST%>%
      filter(lender %in% c('CITI','SBI','Yes Bank','HOME CREDIT','KREDITBEE','Upwards','HDFC','HDFCEMERGINGMARKET','MUTHOOT','IDFC','Bajaj Finserve','Lending Kart','Sme Corner','Sme Corner','FLEXILOAN','INDIFI','Triti','Bajaj','DCB','HFFC','WheelsEMI','HDFC Life','Max Bupa') |
               sku  %in% c('RBL BL') &
               sku  %nin% c('HDFC CC'))%>%
      select(sku)

    SKU_390<-SKU_LENDER_LIST%>%
      filter(lender %in% c('RBL','Early Salary','EARLYSALARYPL','MONEYVIEW','Paysense','CashE') |
               sku  %in% c('HDFC CC') &
               sku  %nin% c('RBL BL'))%>%
      select(sku)

    SKU_270_2<-SKU_LENDER_LIST%>%
      filter(lender %in% c('Shriram'))%>%
      select(sku)

    if(nrow(SKU_270)!=0){
      MOVED_270 <- MTD_FUNNEL_S2L_2[MTD_FUNNEL_S2L_2$SKU %in% SKU_270$sku,]
      MOVED_270<-data.frame('SKU'=MOVED_270[[as.numeric(which(names(MOVED_270)[]=='SKU'))]],
                            'Moved to 270'=MOVED_270[[as.numeric(which(names(MOVED_270)[]=='Moved to 270'))]]
      )
      names(MOVED_270)[2]='Moved to 270'
    }

    if(nrow(SKU_390)!=0){
      MOVED_390 <- MTD_FUNNEL_S2L_2[MTD_FUNNEL_S2L_2$SKU %in% SKU_390$sku,]
      MOVED_390<-data.frame('SKU'=MOVED_390[[as.numeric(which(names(MOVED_390)[]=='SKU'))]],
                            'Moved to 390'=MOVED_390[[as.numeric(which(names(MOVED_390)[]=='Moved to 390'))]]
      )
      names(MOVED_390)[2]='Moved to 390'
    }

    if(nrow(SKU_270_2)!=0){
      MOVED_270_2 <- MTD_FUNNEL_S2L_2[MTD_FUNNEL_S2L_2$SKU %in% SKU_270_2$sku,]
      MOVED_270_2<-data.frame('SKU'=MOVED_270_2[[as.numeric(which(names(MOVED_270_2)[]=='SKU'))]],
                              'MOVED_270'=MOVED_270_2[[as.numeric(which(names(MOVED_270_2)[]=='MOVED_270'))]]
      )
    }

    if(nrow(MOVED_390)!=0 || nrow(MOVED_270)!=0 || nrow(MOVED_270_2)!=0 ){
      names(MOVED_270)<-names(MOVED_390)
      names(MOVED_270_2)<-names(MOVED_390)

      S2L_ACTUALS<-as.data.frame(rbind(MOVED_390,MOVED_270,MOVED_270_2))
      b<-S2L_ACTUALS %>%
        left_join(y = MTD_FUNNEL_S2L_2, by = c('SKU'))

      c<-b%>%select_if(names(.) %in% c('SKU','Product','Applications','Moved to 390.x'))

      c<-c%>%
        mutate('Pass %'=as_perc(round((b$`Moved to 390.x`/b$Applications),0)))%>%
        mutate('Plan' = round(b$`Moved to 390.x` * as.numeric(b$`Moved to 390.x`/b$Applications),0))%>%
        mutate('Backlog' = round((`Plan`- b$`Moved to 390.x`),0))%>%
        mutate('Achieved %' = as_perc(b$`Moved to 390.x`/`Plan`))

      c<-c%>%select_if(names(.) %in% c('SKU','Product','Applications','Moved to 390.x','Pass %','Plan','Backlog','Achieved %'))

      column_key<-c(SKU='A',Product='B','Moved to 390.x'='C',Plan='D','Pass %'='E',Applications='F','Backlog'='G','Achieved %'='H')
      names(c) <- column_key[names(c)]
      c<-c[ , order(names(c))]
      column_key_2<-c(A='SKU',B='Product',C='Actuals',D='Plan',E='Pass %',F='Applications',G='Backlog',H='Achieved %')
      names(c) <- column_key_2[names(c)]
      S2L_final<-c
    }

  }

}else{
  S2L_final<-data.frame('SKU'=character(0),'Product'=character(0),'Actuals'=numeric(0),'Pass %'=numeric(0),'Applications'=numeric(0),'Backlog'=numeric(0),'Achieved %'=numeric(0))
}

COL_START = column_no + 3
write_excel_table(wb, file = "C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\Referrals Forecast.csv", S2L_final, 3, COL_START,"S2L/API Passed")
#rajan changes end 09/08/2021