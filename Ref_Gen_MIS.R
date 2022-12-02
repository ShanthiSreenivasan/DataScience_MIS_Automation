CON <-
  dbConnect(
    PostgreSQL(),
    port = DBCONFIG$warehouse$port,
    user = DBCONFIG$warehouse$username,
    password = DBCONFIG$warehouse$password,
    dbname = DBCONFIG$warehouse$db,
    host = DBCONFIG$warehouse$host
  )

loginfo("Referral Automation - Report initated")

'%nin%' <- Negate(`%in%`)

aws.s3::save_object(
  str_interp("s3://cm-referrals-reports/cm-referrals-reports/07 Dump - ${DATE}.zip"),
  REFERRALS_REPORTS_BUCKET,
  file = str_interp(basename("07 Dump - ${DATE}.zip")),
  headers = list(),
  overwrite = TRUE
)

unzip(file.path(Sys.getenv('CM_REPORTS_FOLDER'), str_interp("07 Dump - ${DATE}.zip")))

data<-read.csv(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("tmp/cm-reports-data/${DATE}/Referrals Dump - ${DATE}.csv")))
month_target <-read.xlsx(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("revised target sheet.xlsx")),sheet=1)
m0_target <-read.xlsx(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("revised target sheet.xlsx")),sheet=2)

pt<-data %>% group_by(sku,lender,product,attribution) %>% summarise(`MTD_Actuals`=n_distinct(lead_id)) %>% data.frame()


pt_1<-spread(pt,attribution,MTD_Actuals)

names(pt_1)<-replace(names(pt_1),match(c('Assisted'),names(pt_1)),'Assisted_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('Marketing'),names(pt_1)),'Marketing_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('PA'),names(pt_1)),'PA_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('Regeneration_Others'),names(pt_1)),'Regeneration_Others_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('Regeneration_PA'),names(pt_1)),'Regeneration_PA_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('Regeneration_STP'),names(pt_1)),'Regeneration_STP_MTD_Actuals')
names(pt_1)<-replace(names(pt_1),match(c('System'),names(pt_1)),'STP_MTD_Actuals')

pt_1$`PA_MTD_Actuals`<-pt_1$Marketing_MTD_Actuals+pt_1$PA_MTD_Actuals
pt_1 <-pt_1 %>% select (-`Marketing_MTD_Actuals`)


if("Regeneration_Others_MTD_Actuals" %in% colnames(pt_1) && "Regeneration_PA_MTD_Actuals" %in% colnames(pt_1) && "Regeneration_STP_MTD_Actuals" %in% colnames(pt_1)){
  pt_1 <- pt_1 %>% mutate (`Regen_MTD_Actuals`= pt_1$Regeneration_Others_MTD_Actuals+pt_1$Regeneration_PA_MTD_Actuals+pt_1$Regeneration_STP_MTD_Actuals)
  pt_1 <-pt_1 %>% select(-`Regeneration_Others_MTD_Actuals`,-`Regeneration_PA_MTD_Actuals`,-`Regeneration_STP_MTD_Actuals`)
}else{pt_1 <- pt_1 %>% mutate (`Regen_MTD_Actuals`= 0)}

pt_1$sku <- toupper(pt_1$sku)
month_target$sku <- toupper(month_target$sku)

pt_targets <- merge(pt_1,month_target,by.x="sku",by.y="sku",all.x=T,all.y=F)
pt_targets <- pt_targets %>% rename(`STP_Month_Targets`=`System`,`PA_Month_Targets`=`PA`,`Assisted_Month_Targets`=`Assisted`,`Networks_Month_Targets`=`Networks`,
                                    `Allset_Month_Targets`=`Allset`,`product`=`product.x`)

pt_targets <-pt_targets %>% select(-`product.y`)

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

pt_targets <- pt_targets %>% mutate(`Assisted DRR`=((pt_targets$Assisted_Month_Targets-pt_targets$Assisted_MTD_Actuals)/assisted_remaining_days))

pa_remaining_days<-pa_target_mtd-DAYS_PAST

if((DAYS_PAST) <=25){
pt_targets <- pt_targets %>% mutate(`STP DRR`=((pt_targets$STP_Month_Targets-pt_targets$STP_MTD_Actuals)/pa_remaining_days))
}else{
pt_targets <- pt_targets %>% mutate(`STP DRR`=((pt_targets$STP_Month_Targets-pt_targets$STP_MTD_Actuals)/(DAYS_LEFT)))
}

if((DAYS_PAST) <=25){
  pt_targets <- pt_targets %>% mutate(`Portfolio DRR`=((pt_targets$PA_Month_Targets-pt_targets$PA_MTD_Actuals)/pa_remaining_days))
}else{
  pt_targets <- pt_targets %>% mutate(`Portfolio DRR`=((pt_targets$PA_Month_Targets-pt_targets$PA_MTD_Actuals)/(DAYS_LEFT)))
}

pt_targets $`Assisted DRR` <-round(pt_targets$`Assisted DRR`,0)
pt_targets $`Portfolio DRR` <-round(pt_targets$`Portfolio DRR`,0)
pt_targets $`STP DRR` <-round(pt_targets$`STP DRR`,0)


pt_targets <- pt_targets %>% mutate(`Assisted MTD Backlog`=(pt_targets$`Assisted MTD Target`-pt_targets$Assisted_MTD_Actuals))

pt_targets <- pt_targets %>% mutate(`STP MTD Backlog`=(pt_targets$`STP MTD Target`-pt_targets$STP_MTD_Actuals))

pt_targets <- pt_targets %>% mutate(`Portfolio MTD Backlog`=(pt_targets$`PA MTD Target`-pt_targets$PA_MTD_Actuals))


pt_targets <- pt_targets %>% mutate(`Assisted % Achieved`=((pt_targets$Assisted_MTD_Actuals)/(pt_targets$`Assisted MTD Target`)*100))

pt_targets <- pt_targets %>% mutate(`STP % Achieved`=((pt_targets$STP_MTD_Actuals)/(pt_targets$`STP MTD Target`)*100))

pt_targets <- pt_targets %>% mutate(`Portfolio % Achieved`=((pt_targets$PA_MTD_Actuals)/(pt_targets$`PA MTD Target`)*100))

pt_targets $`Assisted % Achieved` <-round(pt_targets$`Assisted % Achieved`,0)
pt_targets $`STP % Achieved` <-round(pt_targets$`STP % Achieved`,0)
pt_targets $`Portfolio % Achieved` <-round(pt_targets$`Portfolio % Achieved`,0)

pt_targets$`Assisted % Achieved` <- ifelse(pt_targets$`Assisted % Achieved` %in% c('Inf'),0,pt_targets$`Assisted % Achieved`)
pt_targets$`STP % Achieved` <- ifelse(pt_targets$`STP % Achieved` %in% c('Inf'),0,pt_targets$`STP % Achieved`)
pt_targets$`Portfolio % Achieved` <- ifelse(pt_targets$`Portfolio % Achieved` %in% c('Inf'),0,pt_targets$`Portfolio % Achieved`)

##########################################################################################################################################

send_to_lender_actuals_1 <- data %>% filter(appops_status_code >=270 & !appops_status_code %in% c(370,380,451,461,710) &
                                          sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                     'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                     'L & T Finance PL','L & T Finance SBPL','Smart Coin CL',
                                                     'HOME CREDIT CL','Yes Bank PL','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL')) %>% group_by(sku,appops_status_code) %>% summarise(`S2L Actuals`=n_distinct(lead_id))

send_to_lender_actuals_1 <- spread(send_to_lender_actuals_1,appops_status_code,`S2L Actuals`)

send_to_lender_actuals_1_cnt <-rowSums(send_to_lender_actuals_1[,2:ncol(send_to_lender_actuals_1)],na.rm=TRUE) %>% data.frame()

send_to_lender_actuals_1_sku <- send_to_lender_actuals_1 %>% select(sku) %>% data.frame()

send_to_lender_actuals_1_data <-cbind(send_to_lender_actuals_1_sku,send_to_lender_actuals_1_cnt) %>% data.frame()

names(send_to_lender_actuals_1_data)[2] <-c("S2L/API Actuals")

##########################################################################################################################################
send_to_lender_actuals_2 <- data %>% filter(appops_status_code >=390 & !appops_status_code %in% c(710) &
                                              !sku %in% c('INDIFI BL','CITI CC','PNB HL','PNB HLBT','PNB LAP','India Shelter HL',
                                                         'India Shelter HLBT','India Shelter LAP','MUTHOOT PL','MUTHOOT SBPL',
                                                         'L & T Finance PL','L & T Finance SBPL','Smart Coin CL',
                                                         'HOME CREDIT CL','Yes Bank PL','Yes Bank SBPL','AXIS BANK CC','AXIS BANK PL')) %>% group_by(sku,appops_status_code) %>% summarise(`S2L Actuals`=n_distinct(lead_id))

send_to_lender_actuals_2 <- spread(send_to_lender_actuals_2,appops_status_code,`S2L Actuals`)

send_to_lender_actuals_2_cnt <-rowSums(send_to_lender_actuals_2[,2:ncol(send_to_lender_actuals_2)],na.rm=TRUE) %>% data.frame()

send_to_lender_actuals_2_sku <- send_to_lender_actuals_2 %>% select(sku) %>% data.frame()

send_to_lender_actuals_2_data <-cbind(send_to_lender_actuals_2_sku,send_to_lender_actuals_2_cnt) %>% data.frame()

names(send_to_lender_actuals_2_data)[2] <-c("S2L/API Actuals")


##########################################################################################################################################

send_to_lenders_actuals <-rbind(send_to_lender_actuals_1_data,send_to_lender_actuals_2_data)

send_to_lenders_actuals$sku <- toupper(send_to_lenders_actuals$sku)

final_pt_targets <- list(pt_targets,send_to_lenders_actuals) %>% reduce(left_join,by="sku")

########################################################################################################################################

overall_07 <- data %>% group_by(sku) %>% summarise(`Overall_07`=n_distinct(lead_id))
overall_07$sku <- toupper(overall_07$sku)
final_pt_targets <-list(final_pt_targets,overall_07) %>% reduce(left_join,by="sku")


#########################################################################################################################################

conversions_990 <- data %>% filter(appops_status_code %in% c(990)) %>% group_by(sku) %>%  summarise(`990 conversions`=n_distinct(lead_id))
conversions_990$sku <- toupper(conversions_990$sku)
final_pt_targets <-list(final_pt_targets,conversions_990) %>% reduce(left_join,by="sku")


network_mtd_actuals <- data %>% filter(profile_vintage %in% c('M0') & nw_flag ==1) %>% group_by(sku) %>% summarise(`Networks MTD Actuals`=n_distinct(lead_id))
network_mtd_actuals$sku <- toupper(network_mtd_actuals$sku)

alliance_mtd_actuals <- data %>% filter(profile_vintage %in% c('M0') & alliance_flag ==1) %>% group_by(sku) %>% summarise(`Alliance MTD Actuals`=n_distinct(lead_id))
alliance_mtd_actuals$sku <- toupper(alliance_mtd_actuals$sku)

final_pt_targets <-list(final_pt_targets,network_mtd_actuals,alliance_mtd_actuals) %>% reduce(left_join,by="sku")

###########################################################################################################################################

final_pt_targets$new_lender <-ifelse(grepl('cash',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('STBL'),'CASHE (STBL,BNPL)',
                              ifelse(grepl('cash',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('BNPL'),'CASHE (STBL,BNPL)',
                              ifelse(grepl('early',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('BNPL'),'EARLYSALARY (STBL,BNPL,PL,SBPL)',
                              ifelse(grepl('early',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('STBL'),'EARLYSALARY (STBL,BNPL,PL,SBPL)',
                              ifelse(grepl('early',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'EARLYSALARY (STBL,BNPL,PL,SBPL)',
                              ifelse(grepl('early',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'EARLYSALARY (STBL,BNPL,PL,SBPL)',
                              ifelse(grepl('home',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('CL'),'HOME CREDIT (CL,STBL,PL,SBPL)',
                              ifelse(grepl('home',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('STBL'),'HOME CREDIT (CL,STBL,PL,SBPL)',
                              ifelse(grepl('home',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'HOME CREDIT (CL,STBL,PL,SBPL)',
                              ifelse(grepl('home',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'HOME CREDIT (CL,STBL,PL,SBPL)',       
                              ifelse(grepl('kredit',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('CL'),'KREDITBEE (CL,STBL,CC,PL)',
                              ifelse(grepl('kredit',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('STBL'),'KREDITBEE (CL,STBL,CC,PL)',
                              ifelse(grepl('kredit',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('CC'),'KREDITBEE (CL,STBL,CC,PL)',
                              ifelse(grepl('kredit',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'KREDITBEE (CL,STBL,CC,PL)',
                              ifelse(grepl('smart',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('CL'),'SMARTCOIN (CL,STBL)',
                              ifelse(grepl('smart',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('STBL'),'SMARTCOIN (CL,STBL)',
                              ifelse(grepl('shriram',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'Shriram (PL,SBPL)',
                              ifelse(grepl('shriram',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'Shriram (PL,SBPL)',
                              ifelse(grepl('HDFC',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'HDFC (PL,SBPL)',
                              ifelse(grepl('HDFC',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'HDFC (PL,SBPL)',
                              ifelse(grepl('faircent',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'Faircent (PL,SBPL)',
                              ifelse(grepl('faircent',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'Faircent (PL,SBPL)',
                              ifelse(grepl('muthoot',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'Muthoot (PL,SBPL)',
                              ifelse(grepl('muthoot',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'Muthoot (PL,SBPL)',
                              ifelse(grepl('axis',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'AXIS (PL,SBPL)',
                              ifelse(grepl('axis',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'AXIS (PL,SBPL)',
                              ifelse(grepl('finzy',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'Finzy (PL,SBPL)',
                              ifelse(grepl('finzy',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'Finzy (PL,SBPL)',
                              ifelse(grepl('IDFC',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'IDFC (PL,SBPL)',
                              ifelse(grepl('IDFC',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'IDFC (PL,SBPL)',
                              ifelse(grepl('YES',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'YesBank (PL,SBPL)',
                              ifelse(grepl('YES',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'YesBank (PL,SBPL)',
                              ifelse(grepl('L & T Finance',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('PL'),'L & T Finance (PL,SBPL)',
                              ifelse(grepl('L & T Finance',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('SBPL'),'L & T Finance (PL,SBPL)',
                              ifelse(grepl('PNB',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HL'),'PNB(HL,HLBT,LAP)',
                              ifelse(grepl('PNB',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HLBT'),'PNB(HL,HLBT,LAP)',
                              ifelse(grepl('PNB',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('LAP'),'PNB(HL,HLBT,LAP)',                                      
                              ifelse(grepl('India Shelter',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HL'),'India Shelter(HL,HLBT,LAP)',
                              ifelse(grepl('India Shelter',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HLBT'),'India Shelter(HL,HLBT,LAP)',
                              ifelse(grepl('India Shelter',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('LAP'),'India Shelter(HL,HLBT,LAP)',    
                              ifelse(grepl('canara',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HL'),'Canara(HL,HLBT,LAP)',
                              ifelse(grepl('canara',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('HLBT'),'Canara(HL,HLBT,LAP)',
                              ifelse(grepl('canara',final_pt_targets$lender,ignore.case = T) & final_pt_targets$product %in% c('LAP'),'Canara(HL,HLBT,LAP)',
                              final_pt_targets$sku)))))))))))))))))))))))))))))))))))))))))))


final_ref_gen <- final_pt_targets %>% group_by(new_lender) %>% summarise(`Assisted_MTD_Actuals`=sum(Assisted_MTD_Actuals,na.rm = TRUE),
                                                                         `PA_MTD_Actuals`=sum(PA_MTD_Actuals,na.rm = TRUE),
                                                                         `STP_MTD_Actuals`=sum(STP_MTD_Actuals,na.rm = TRUE),
                                                                         `Regen_MTD_Actuals`=sum(Regen_MTD_Actuals,na.rm = TRUE),
                                                                         `STP_Month_Targets`=sum(STP_Month_Targets,na.rm = TRUE),
                                                                         `PA_Month_Targets`=sum(PA_Month_Targets,na.rm = TRUE),
                                                                         `Assisted_Month_Targets`=sum(Assisted_Month_Targets,na.rm = TRUE),
                                                                         `Networks_Month_Targets`=sum(Networks_Month_Targets,na.rm = TRUE),
                                                                         `Allset_Month_Targets`=sum(Allset_Month_Targets,na.rm = TRUE),
                                                                         `Assisted MTD Target`=sum(`Assisted MTD Target`,na.rm = TRUE),
                                                                         `PA MTD Target`=sum(`PA MTD Target`,na.rm = TRUE),
                                                                         `STP MTD Target`=sum(`STP MTD Target`,na.rm = TRUE),
                                                                         `Assisted DRR`=sum(`Assisted DRR`,na.rm = TRUE),
                                                                         `STP DRR`=sum(`STP DRR`,na.rm = TRUE),
                                                                         `Portfolio DRR`=sum(`Portfolio DRR`,na.rm = TRUE),
                                                                         `Assisted MTD Backlog`=sum(`Assisted MTD Backlog`,na.rm = TRUE),
                                                                         `STP MTD Backlog`=sum(`STP MTD Backlog`,na.rm = TRUE),
                                                                         `Portfolio MTD Backlog`=sum(`Portfolio MTD Backlog`,na.rm = TRUE),
                                                                         `Assisted % Achieved`=sum(`Assisted % Achieved`,na.rm = TRUE),
                                                                         `STP % Achieved`=sum(`STP % Achieved`,na.rm = TRUE),
                                                                         `Portfolio % Achieved`=sum(`Portfolio % Achieved`,na.rm = TRUE),
                                                                         `S2L/API Actuals`=sum(`S2L/API Actuals`,na.rm = TRUE),
                                                                         `Overall_07`=sum(`Overall_07`,na.rm = TRUE),
                                                                         `990 conversions`=sum(`990 conversions`,na.rm = TRUE),
                                                                         `Networks MTD Actuals`=sum(`Networks MTD Actuals`,na.rm = TRUE),
                                                                         `Alliance MTD Actuals`=sum(`Alliance MTD Actuals`,na.rm = TRUE))

final_ref_gen <- final_ref_gen %>% mutate(`S2L/API passed %`=((final_ref_gen$`S2L/API Actuals`)/(final_ref_gen$`Overall_07`)*100))
final_ref_gen$`S2L/API passed %` <- round(final_ref_gen$`S2L/API passed %`,0)

final_ref_gen <- final_ref_gen %>% mutate(`S2L/API Plan`=(final_ref_gen$`Overall_07`*final_ref_gen$`S2L/API passed %`)/100)
final_ref_gen$`S2L/API Plan` <- round(final_ref_gen$`S2L/API Plan`,0)

final_ref_gen <- final_ref_gen %>% mutate(`S2L/API Backlog`=(final_ref_gen$`S2L/API Plan`- final_ref_gen$`S2L/API Actuals`))

final_ref_gen <- final_ref_gen %>% mutate(`S2L/API % Achieved`=((final_ref_gen$`S2L/API Actuals`)/(final_ref_gen$`S2L/API Plan`)*100))
final_ref_gen$`S2L/API % Achieved` <- round(final_ref_gen$`S2L/API % Achieved`,0)
final_ref_gen$`S2L/API % Achieved` <- ifelse(final_ref_gen$`S2L/API % Achieved` %in% c('Inf'),0,final_ref_gen$`S2L/API % Achieved`)


final_ref_gen <- final_ref_gen %>% mutate(`Conversions %`=((final_ref_gen$`990 conversions`)/(final_ref_gen$`S2L/API Actuals`)*100))
final_ref_gen$`Conversions %` <- round(final_ref_gen$`Conversions %`,0)                                                                         

names(final_ref_gen)[1]<-c("SKU")

final_ref_gen <- final_ref_gen %>% mutate(`MTD Actuals W/O Regen`=final_ref_gen$Assisted_MTD_Actuals+final_ref_gen$PA_MTD_Actuals+final_ref_gen$STP_MTD_Actuals)

final_ref_gen <- final_ref_gen %>% mutate(`MTD Actuals W Regen`=final_ref_gen$Assisted_MTD_Actuals+final_ref_gen$PA_MTD_Actuals+final_ref_gen$STP_MTD_Actuals+final_ref_gen$Regen_MTD_Actuals)

final_ref_gen <- final_ref_gen %>% mutate(`MTD Plan`=final_ref_gen$Assisted_Month_Targets+final_ref_gen$PA_Month_Targets+final_ref_gen$STP_Month_Targets)

final_ref_gen <- final_ref_gen %>% mutate(`MTD Backlog`=final_ref_gen$`MTD Plan`-final_ref_gen$`MTD Actuals W/O Regen`)

final_ref_gen <- final_ref_gen %>% mutate(`% Achieved`=((final_ref_gen$`MTD Actuals W/O Regen`)/(final_ref_gen$`MTD Plan`)*100))

final_ref_gen$`% Achieved` <- ifelse(final_ref_gen $`% Achieved` %in% c('Inf'),0,final_ref_gen $`% Achieved`)
final_ref_gen $`% Achieved` <-round(final_ref_gen $`% Achieved`,0)
final_ref_gen <- final_ref_gen %>% select (SKU,`MTD Actuals W/O Regen`,`MTD Actuals W Regen`,`MTD Plan`,`MTD Backlog`,`% Achieved`,
                                           Assisted_Month_Targets,`Assisted DRR`,Assisted_MTD_Actuals,`Assisted MTD Target`,`Assisted MTD Backlog`,`Assisted % Achieved`,
                                           STP_Month_Targets,`STP DRR`,STP_MTD_Actuals,`STP MTD Target`,`STP MTD Backlog`,`STP % Achieved`,
                                           PA_Month_Targets,`Portfolio DRR`,PA_MTD_Actuals,`PA MTD Target`,`Portfolio MTD Backlog`,`Portfolio % Achieved`,
                                           Networks_Month_Targets,`Networks MTD Actuals`,`Alliance MTD Actuals`,Regen_MTD_Actuals,`S2L/API Actuals`,`S2L/API Plan`,
                                           `S2L/API passed %`,`S2L/API Backlog`,`S2L/API % Achieved`,`990 conversions`,`Conversions %`)

final_ref_gen_total <- colSums(final_ref_gen[,2:ncol(final_ref_gen)],na.rm=TRUE)
final_ref<-bind_rows(final_ref_gen,final_ref_gen_total)
final_ref[nrow(final_ref),1]<-'Total'


# file.remove(file.path(Sys.getenv('CM_REPORTS_FOLDER'), str_interp("07 Dump - ${DATE}.zip")))
# file.remove(file.path(Sys.getenv('CM_REPORTS_FOLDER'),str_interp("tmp/cm-reports-data/${DATE}/Referrals Dump - ${DATE}.csv")))


write_excel_table(wb,'REF',final_ref,3,3,'Referrals Generated 07 Leads Summary')