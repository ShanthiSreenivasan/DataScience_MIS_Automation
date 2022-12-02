CON <-
  dbConnect(
    PostgreSQL(),
    port = DBCONFIG$warehouse$port,
    user = DBCONFIG$warehouse$username,
    password = DBCONFIG$warehouse$password,
    dbname = DBCONFIG$warehouse$db,
    host = DBCONFIG$warehouse$host
  )

loginfo("Referral Automation -STUCK CASES Report initated")

data<-read.csv(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("tmp/cm-reports-data/${DATE}/Referrals Dump - ${DATE}.csv")))
month_target <-read.xlsx(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("revised target sheet.xlsx")),sheet=1)
m0_target <-read.xlsx(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("revised target sheet.xlsx")),sheet=2)

loginfo("Referral Automation -STUCK CASES Querying Done")

pt<-data %>% group_by(product) %>% summarise(`MTD_Actuals`=n_distinct(lead_id))
month_target <- month_target %>% select(product,System,PA,Assisted)
month_target_pro <- month_target %>% group_by(product) %>% summarise(`System`=sum(System,na.rm = TRUE),
                                                                     `PA`=sum(PA,na.rm = TRUE),
                                                                     `Assisted`=sum(Assisted,na.rm = TRUE))

pt_targets <- merge(pt,month_target_pro,by.x="product",by.y="product",all.x=T,all.y=F)
pt_targets <- pt_targets %>% rename(`STP_Month_Targets`=`System`,`PA_Month_Targets`=`PA`,`Assisted_Month_Targets`=`Assisted`)


assisted_total_days <-as.numeric(25)
assisted_off_days <-as.numeric(0)
assisted_working_days<- DAYS_PAST-assisted_off_days
assisted_remaining_days <-assisted_total_days-assisted_working_days

pt_targets <- pt_targets %>% mutate(`Assisted MTD Target`= ((pt_targets$Assisted_Month_Targets)/(assisted_total_days))*(assisted_working_days))

pa_target_mtd <-as.numeric(25)
pa_target_full <-DAYS_IN_MONTH

if((DAYS_PAST) <= 25){
  pt_targets <- pt_targets %>% mutate(`PA MTD Target`=((pt_targets$PA_Month_Targets)/(pa_target_mtd))*(DAYS_PAST))
}else{
  pt_targets <- pt_targets %>% mutate(`PA MTD Target`=((pt_targets$PA_Month_Targets)/(pa_target_full))*(DAYS_PAST))
}

if((DAYS_PAST) <= 25){
  pt_targets <- pt_targets %>% mutate(`STP MTD Target`=((pt_targets$STP_Month_Targets)/(pa_target_mtd))*(DAYS_PAST))
}else{
  pt_targets <- pt_targets %>% mutate(`STP MTD Target`=((pt_targets$STP_Month_Targets)/(pa_target_full))*(DAYS_PAST))
}

pt_targets $`Assisted MTD Target` <-round(pt_targets$`Assisted MTD Target`,0)
pt_targets $`PA MTD Target` <-round(pt_targets$`PA MTD Target`,0)
pt_targets $`STP MTD Target` <-round(pt_targets$`STP MTD Target`,0)

assisted_drr<-assisted_total_days-DAYS_PAST

pa_remaining_days<-pa_target_mtd-DAYS_PAST

pt_targets<- pt_targets %>% mutate(`Month_Target`=pt_targets$STP_Month_Targets+pt_targets$PA_Month_Targets+pt_targets$Assisted_Month_Targets)

pt_targets<- pt_targets %>% mutate(`MTD_Target`=pt_targets$`Assisted MTD Target`+pt_targets$`PA MTD Target`+pt_targets$`STP MTD Target`)



##########################################################################################################################################

send_to_lender_actuals_1 <- data %>% filter(appops_status_code >=270 & !appops_status_code %in% c(370,380,451,461,710) &
                                              sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                         'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                         'L & T Finance PL','L & T Finance SBPL','Smart Coin CL',
                                                         'HOME CREDIT CL','Yes Bank PL','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL')) %>% group_by(product,appops_status_code) %>% summarise(`S2L Actuals`=n_distinct(lead_id))

send_to_lender_actuals_1 <- spread(send_to_lender_actuals_1,appops_status_code,`S2L Actuals`)

send_to_lender_actuals_1_cnt <-rowSums(send_to_lender_actuals_1[,2:ncol(send_to_lender_actuals_1)],na.rm=TRUE) %>% data.frame()

send_to_lender_actuals_1_product <- send_to_lender_actuals_1 %>% select(product) %>% data.frame()

send_to_lender_actuals_1_data <-cbind(send_to_lender_actuals_1_product,send_to_lender_actuals_1_cnt) %>% data.frame()

names(send_to_lender_actuals_1_data)[2] <-c("S2L/API Actuals")

##########################################################################################################################################
send_to_lender_actuals_2 <- data %>% filter(appops_status_code >=390 & !appops_status_code %in% c(710) &
                                              !sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                          'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                          'L & T Finance PL','L & T Finance SBPL','Smart Coin CL',
                                                          'HOME CREDIT CL','Yes Bank PL','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL')) %>% group_by(product,appops_status_code) %>% summarise(`S2L Actuals`=n_distinct(lead_id))

send_to_lender_actuals_2 <- spread(send_to_lender_actuals_2,appops_status_code,`S2L Actuals`)

send_to_lender_actuals_2_cnt <-rowSums(send_to_lender_actuals_2[,2:ncol(send_to_lender_actuals_2)],na.rm=TRUE) %>% data.frame()

send_to_lender_actuals_2_product <- send_to_lender_actuals_2 %>% select(product) %>% data.frame()

send_to_lender_actuals_2_data <-cbind(send_to_lender_actuals_2_product,send_to_lender_actuals_2_cnt) %>% data.frame()

names(send_to_lender_actuals_2_data)[2] <-c("S2L/API Actuals")


##########################################################################################################################################

send_to_lenders_actuals <-rbind(send_to_lender_actuals_1_data,send_to_lender_actuals_2_data)

send_to_lenders_actuals <- send_to_lenders_actuals %>% group_by(product) %>% summarise(`S2L/API Actuals`=sum(`S2L/API Actuals`))

final_pt_targets <- list(pt_targets,send_to_lenders_actuals) %>% reduce(left_join,by="product")

########################################################################################################################################

final_pt_targets<- final_pt_targets %>% mutate(`API Passed %`=((final_pt_targets$`S2L/API Actuals`)/(final_pt_targets$MTD_Actuals)*100))

final_pt_targets$`API Passed %` <- round(final_pt_targets$`API Passed %`,0)


stuck_app_ops <- data %>% filter(appops_status_code %in% c(89,140,150,210,240,261,710)) %>% group_by(product,appops_status_code) %>% summarise(`MTD_Actuals`=n_distinct(lead_id)) %>% data.frame()

stuck_app_ops <- spread(stuck_app_ops,appops_status_code,MTD_Actuals)

final_pt_targets <-list(final_pt_targets,stuck_app_ops) %>% reduce(left_join,by="product")
#########################################################################################################################################


stuck_app_ops_total<-rowSums(final_pt_targets[,13:ncol(final_pt_targets)],na.rm=TRUE) %>% data.frame()
names(stuck_app_ops_total)[1] <-c("Stuck Leads")

final_pt_targets<- cbind(final_pt_targets,stuck_app_ops_total)


final_pt_targets <- final_pt_targets %>% mutate(`Stuck %`= ((final_pt_targets$`Stuck Leads`)/(final_pt_targets$MTD_Actuals)*100))
final_pt_targets$`Stuck %` <- round(final_pt_targets$`Stuck %`,0)

final_pt_targets_total <- colSums(final_pt_targets[,2:ncol(final_pt_targets)],na.rm=TRUE)
final_pt_targets<-bind_rows(final_pt_targets,final_pt_targets_total)
final_pt_targets[nrow(final_pt_targets),1]<-'Total'

final_pt_targets <- final_pt_targets %>% select(-`STP_Month_Targets`,-`PA_Month_Targets`,-`Assisted_Month_Targets`,-`Assisted MTD Target`,-`PA MTD Target`,-`STP MTD Target`)
names(final_pt_targets)

##########################################################################################################################################


s2l_api_stuck_app_ops <- data %>% filter(appops_status_code %in% c(140,150,210,240,261,710) & 
                                           !sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                      'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                      'L & T Finance PL','L & T Finance SBPL','Smart Coin CL','Yes Bank CL',
                                                      'HOME CREDIT CL','Yes Bank PL','Yes Bank CC','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL'))%>% 
                                                       group_by(product,lender,appops_status_code) %>% summarise(`MTD_Actuals`=n_distinct(lead_id)) %>% 
                                                       data.frame()

s2l_api_stuck_app_ops <- spread(s2l_api_stuck_app_ops,appops_status_code,MTD_Actuals)

s2l_api_stuck_app_ops_cnt <-rowSums(s2l_api_stuck_app_ops[,3:ncol(s2l_api_stuck_app_ops)],na.rm=TRUE) %>% data.frame()

names(s2l_api_stuck_app_ops_cnt)[1] <-c("Total")

s2l_api_stuck_app_ops_data <-cbind(s2l_api_stuck_app_ops,s2l_api_stuck_app_ops_cnt)

s2l_api_stuck_app_ops_data_total <- colSums(s2l_api_stuck_app_ops_data[,3:ncol(s2l_api_stuck_app_ops_data)],na.rm=TRUE)
final_ref_s2l_api<-bind_rows(s2l_api_stuck_app_ops_data,s2l_api_stuck_app_ops_data_total)
final_ref_s2l_api[nrow(final_ref_s2l_api),1]<-'Total'

#############################################################################################################################################

manual_api_stuck_app_ops <- data %>% filter(appops_status_code %in% c(140,150,210,240,261,710) & 
                                            sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                       'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                       'L & T Finance PL','L & T Finance SBPL','Smart Coin CL','Yes Bank CL',
                                                       'HOME CREDIT CL','Yes Bank PL','Yes Bank CC','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL'))%>% 
                                                        group_by(product,lender,appops_status_code) %>% summarise(`MTD_Actuals`=n_distinct(lead_id)) %>% 
                                                        data.frame()

manual_api_stuck_app_ops <- spread(manual_api_stuck_app_ops,appops_status_code,MTD_Actuals)

manual_api_stuck_app_ops_cnt <-rowSums(manual_api_stuck_app_ops[,3:ncol(manual_api_stuck_app_ops)],na.rm=TRUE) %>% data.frame()

names(manual_api_stuck_app_ops_cnt)[1] <-c("Total")

manual_api_stuck_app_ops_data <-cbind(manual_api_stuck_app_ops,manual_api_stuck_app_ops_cnt)

manual_api_stuck_app_ops_data_total <- colSums(manual_api_stuck_app_ops_data[,3:ncol(manual_api_stuck_app_ops_data)],na.rm=TRUE)
final_ref_manual_api<-bind_rows(manual_api_stuck_app_ops_data,manual_api_stuck_app_ops_data_total)
final_ref_manual_api[nrow(final_ref_manual_api),1]<-'Total'

#############################################################################################################################################

write_excel_table(wb,'STUCK CASES',final_pt_targets,3,3,'MTD')
write_excel_table(wb,'STUCK CASES',final_ref_s2l_api,3+nrow(final_pt_targets)+3,3,'S2L thru API Stuck Leads Summary')
write_excel_table(wb,'STUCK CASES',final_ref_manual_api,3+nrow(final_pt_targets)+3,3+ncol(final_ref_s2l_api)+3,'(Manual S2L) Leads Summary')

loginfo("Referral Automation - Written to Excel")
loginfo("Referral Automation - STUCK CASES Done")

 
file.remove(file.path(Sys.getenv('CM_REPORTS_FOLDER'), str_interp("07 Dump - ${DATE}.zip")))
file.remove(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("tmp/cm-reports-data/${DATE}/Referrals Dump - ${DATE}.csv")))