rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)



setwd('C:\\R\\CIS_DIALER OPS Automation')

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

CISdump <- fread('./Input/CIS_NEW_V0_DIALER_BASE.csv')

CISdump$user_id <- as.numeric(CISdump$user_id)

CISdump$phone_home <- as.numeric(CISdump$phone_home)

CISdump$customer_name <- str_c(CISdump$first_name,'',CISdump$last_name)

CISdump$latest_login <- as.Date(CISdump$latest_login,             # Change class of date column
                                    format = "%d-%m-%Y")
CISdump <- CISdump[rev(order(CISdump$latest_login)), ]


Output_fields <- read_excel(".\\Input\\Automation_CIS_Dialer.xlsx",sheet='Sheet1')

#Output_fields <- Output_fields %>% filter(files =='1')
  
'%nin%' <- Negate(`%in%`)

for(i in Output_fields$Language){
  
  Base_list <- Output_fields %>% filter(Language %in% c(i)) 

  negative_status_accounts_1 <- lapply(strsplit(Base_list$negative_status_accounts,","),as.character)
  
  negative_status_accounts_1 <- unlist(negative_status_accounts_1)
  
  latest_ps_1  <- lapply(strsplit(Base_list$latest_ps,","),as.character)
  
  latest_ps_1 <- unlist(latest_ps_1)
  
  nsaleable_1 <- Base_list$nsaleable
  
  Language_1  <- lapply(strsplit(Base_list$Language,","),as.character)
  
  Language_1 <- unlist(Language_1)
  
  latest_dispo_1  <- lapply(strsplit(Base_list$latest_dispo,","),as.character)
  
  latest_dispo_1 <- unlist(latest_dispo_1)
  
  
  camp <- Base_list$Language
  
  Base_dump <- CISdump %>% filter(Base %in% c(i))
  
  Base_dump_1 <- Base_dump %>% filter(negative_status_accounts %in% negative_status_accounts_1)
  
  Base_dump_2 <- Base_dump_1 %>% filter(latest_ps %in% latest_ps_1)
  
  Base_dump_3 <- Base_dump_2 %>% filter(nsaleable %nin% nsaleable_1)
  
  Base_dump_4 <- Base_dump_3 %>% filter(latest_dispo %nin% latest_dispo_1)
    if(!is.na(Language_1[1]))
  {
    
  Base_dump_5 <- Base_dump_4 %>% filter(Language %in% Language_1)
  }else{
  Base_dump_5 <- Base_dump_4
  }
  Base_dump_5$Language <- camp
  
  #Base_dump_5$base_type <- base_1
  
  Base_dump_6 <- Base_dump_5 %>% select(user_id,customer_name,phone_home,Language) 
  
  Filename <- paste('.\\Output\\',thisdate,'\\',i,"_COMNET_",camp,".csv",sep = '')
  FileCreate(Base_dump_6,sheet_name="Sheet1",Filename)
  
}
 
