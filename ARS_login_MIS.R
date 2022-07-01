rm(list=ls())
library(magrittr)
library(tibble)
library(dplyr)
library(plyr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
library(openxlsx)
library(purrr)
library(janitor)
library(scales)
library(tidyverse)
require(data.table)
require(reshape2)
library(XLConnect)


ars_login_dump <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\Login Tracker_ARS_input.xlsx",sheet = 'ARS Logins Metrics')
lender_login_dump <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\Login Tracker_ARS_input.xlsx",sheet = 'Lender Logins Metrics')

#######################CM Base Total Login
paid = data.frame(ars_login_dump$Paid_sub) %>% summarise(ars_login_dump.Paid_sub=last(ars_login_dump.Paid_sub)) %>% dplyr::rename(`Paid Subs` = ars_login_dump.Paid_sub)
names(CM_base_login)

free = data.frame(ars_login_dump$Free_sub) %>% summarise(ars_login_dump.Free_sub=last(ars_login_dump.Free_sub)) %>% dplyr::rename(`Free_sub` = ars_login_dump.Free_sub)

non = data.frame(ars_login_dump$Profiled.Allocated) %>% summarise(ars_login_dump.Profiled.Allocated=last(ars_login_dump.Profiled.Allocated)) %>% dplyr::rename(`Non-Subs` = ars_login_dump.Profiled.Allocated)

lb = data.frame(lender_login_dump$Profiled.Allocated) %>% summarise(lender_login_dump.Profiled.Allocated=last(lender_login_dump.Profiled.Allocated)) %>% dplyr::rename(`Lender Base` = lender_login_dump.Profiled.Allocated)

final_allo<-cbind('Total Login'="Allocated count",paid,free,non,lb)


paid2 = data.frame(ars_login_dump$Login.on.Paid_sub) %>% summarise(ars_login_dump.Login.on.Paid_sub=last(ars_login_dump.Login.on.Paid_sub)) %>% dplyr::rename(`Paid Subs` = ars_login_dump.Login.on.Paid_sub)

free2 = data.frame(ars_login_dump$Login.on.Free_sub) %>% summarise(ars_login_dump.Login.on.Free_sub=last(ars_login_dump.Login.on.Free_sub)) %>% dplyr::rename(`Free_sub` = ars_login_dump.Login.on.Free_sub)

non2 = data.frame(ars_login_dump$Other.Login) %>% summarise(ars_login_dump.Other.Login=last(ars_login_dump.Other.Login)) %>% dplyr::rename(`Non-Subs` = ars_login_dump.Other.Login)

lb2 = data.frame(lender_login_dump$MTD.Login) %>% summarise(lender_login_dump.MTD.Login=last(lender_login_dump.MTD.Login)) %>% dplyr::rename(`Lender Base` = lender_login_dump.MTD.Login)
final_allo1<-cbind('Total Login'="Total Login MTD",paid2,free2,non2,lb2)

#final_allo2<-cbind(arsce_login,pa_login,Mont_login,andr_login,other_login,lb3)

final_tab1 <- bind_rows(final_allo,final_allo1)

#######################Sources of login:
arsce_login = data.frame(ars_login_dump$ARSCE.Login) %>% summarise(ars_login_dump.ARSCE.Login=last(ars_login_dump.ARSCE.Login)) %>% dplyr::rename(`ARSCE Login` = ars_login_dump.ARSCE.Login)

pa_login = data.frame(ars_login_dump$PA.Login) %>% summarise(ars_login_dump.PA.Login=last(ars_login_dump.PA.Login)) %>% dplyr::rename(`PA login` = ars_login_dump.PA.Login)

Mont_login = data.frame(ars_login_dump$Monitoring.Login) %>% summarise(ars_login_dump.Monitoring.Login=last(ars_login_dump.Monitoring.Login)) %>% dplyr::rename(`Monitoring Login` = ars_login_dump.Monitoring.Login)

andr_login = data.frame(ars_login_dump$ANDROID.Login) %>% summarise(ars_login_dump.ANDROID.Login=last(ars_login_dump.ANDROID.Login)) %>% dplyr::rename(`ANDROID Login` = ars_login_dump.ANDROID.Login)

other_login = data.frame(ars_login_dump$Other.UTM.Login) %>% summarise(ars_login_dump.Other.UTM.Login=last(ars_login_dump.Other.UTM.Login)) %>% dplyr::rename(`Other UTM Login` = ars_login_dump.Other.UTM.Login)

lb3 = lb2

final_allo2<-cbind(arsce_login,pa_login,Mont_login,andr_login,other_login,lb3)



#######################Lender wise login target Vs actuals:
login_TARGET <- read.xlsx("C:\\Users\\User\\Desktop\\MIS\\ARS\\Login Tracker_ARS_input.xlsx",sheet = 'Login_target')


CM_base_login <- ars_login_dump %>%  filter(Account.Status %in% c("Total") & product %in% c("Total")) %>%
  group_by(Lender.Name) %>% dplyr::summarise(Total = sum(MTD.Login, na.rm = TRUE)) %>% dplyr::rename(`CM Base Login` = Total) %>% as.data.frame()
CM_base_login <- replace(CM_base_login, is.na(CM_base_login), 0)
LB_base_login <- lender_login_dump %>%  filter(Account.Status %in% c("Total") & product %in% c("Total")) %>%
  group_by(Lender.Name) %>% dplyr::summarise(Total = sum(MTD.Login, na.rm = TRUE)) %>% dplyr::rename(`Lender Base Login` = Total) %>% as.data.frame()
LB_base_login <- replace(LB_base_login, is.na(LB_base_login), 0)

total_days <-as.numeric(days_in_month(today()-1))
DAYS_PAST <- as.numeric(day(Sys.Date()-1))
tar_df1 <- login_TARGET %>% group_by(Lender.Name) %>% as.data.frame() %>% dplyr::rename(`Login Target` = Login.Target) %>% mutate(`MTD Target` = round((`Login Target`/total_days)*DAYS_PAST))


jointdf<- list(tar_df1,CM_base_login,LB_base_login) %>% reduce(full_join,by=c("Lender.Name"))

jointdf <- replace(jointdf, is.na(jointdf), 0)

jointdf = jointdf %>% mutate(`MTD Actuals (CM + Lender base)` = rowSums(jointdf[, c(4:5)]))#, `MTD Target Achievement %` = round(`MTD Actuals (CM + Lender base)`/tar_df1)) %>% as.data.frame()

df = filter_at(jointdf, vars(`Login Target`), all_vars((.) != 0)) %>% as.data.frame() %>% adorn_totals("row")

df <- df %>% mutate(`MTD Target Achievement %` = df[, 7] <- percent((df[, 6] / df[, 3]), accuracy = 0.01))

#write_csv(df, file = "C:\\Users\\User\\Desktop\\CFU_ARS_DSA\\ARS\\Login Tracker_ARS_output.csv")

##########################CM Base STP Login

CM_base_tot<- paid+free+non

tot_targ= login_TARGET %>% dplyr::summarise(Total = sum(Login.Target, na.rm = TRUE))
STP_tar = (tot_targ/2)
STP_mtd_tar = round((STP_tar/total_days)*DAYS_PAST)
STP_mtd_actual = sum(arsce_login,Mont_login,andr_login,other_login)
STP_ach= round(STP_mtd_actual/STP_mtd_tar*100)
STP_log=(STP_mtd_actual/CM_base_tot*100)

final_tab2 <- bind_rows('CM Base Total Allocated' = CM_base_tot,'STP Login Target' =STP_tar,'STP Login Target MTD'=STP_mtd_tar,'STP Login MTD Actuals'=STP_mtd_actual,
                        'STP Login MTD Target Achievement%'=STP_ach, 'STP Login % on the allocation'=STP_log)
##########################CM Base Total Login (STP+PA)
PA_mtd_tar = round((tot_targ/total_days)*DAYS_PAST)
PA_mtd_actual = sum(paid2,free2,non2,lb2)
PA_ach= round(PA_mtd_actual/PA_mtd_tar*100)
PA_log=(PA_mtd_actual/CM_base_tot*100)

final_tab3<- bind_rows('CM Base Total Login (STP+PA)' = CM_base_tot,'Total Login Target'=tot_targ,
                       'Total Login Target MTD' = PA_mtd_tar,'Total Login MTD Actuals'=PA_mtd_actual,
                       'Total Login Target Achievement%' = PA_ach,'Total Login % on the allocation'=PA_log)
list_of_datasets <- list("Name of DataSheet1" = final_tab1, "Name of Datasheet2" = df)
write.xlsx(list_of_datasets, file = "C:\\Users\\User\\Desktop\\MIS\\ARS\\Login Tracker_ARS_output.xlsx")

