rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
library(DBI)
library(RPostgreSQL)
library(RMySQL)
library(logging)
library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)
library(ggrepel)
library(lookup)
library(fs)
library(htmlTable)
#install.packages("htmlTable")

#install.packages("ggrepel")
#install.packages("lookup")

setwd("C:\\R\\Dialer OPS Automation")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

Arsdump <- fread('./Input/CM_REGISTER_BASE.csv')
Arsdump$lead_id <- as.numeric(Arsdump$lead_id)

Arsdump <- Arsdump %>% filter(!is.na(lender))

# Arsdump <- Arsdump %>%
#   mutate(lender = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & grepl('CC|CCC|credit card', product, ignore.case = T) ~ 'HDFC CC', 
#                             TRUE ~ as.character(lender)),
#          lender = case_when(grepl('HDFC|HDFC Bank', lender, ignore.case = T) & !grepl('CC|CCC|credit card', product, ignore.case = T) ~ 'HDFC Retail', 
#                             TRUE ~ as.character(lender)), 
#          lender = lender)
# Arsdump <- Arsdump %>% mutate('Base_Name' = paste0(Arsdump$base_type,'_',Arsdump$dialer_base,'_',Sys.Date(), '_','Woff', '_', Arsdump$lender), 
#                               'Base_Type' = 'Outbound')
#write.csv(Arsdump, file = paste0("C:\\R\\Dialer OPS Automation\\OB_Sub_12314_4_last4_login_woff_IDFC_", Sys.Date(), '.csv'))
  
Output_fields <- read.xlsx(".\\Input\\ARS_CMreg_Automation_input.xlsx",sheet = 'Sheet1')

fields_name <- read.xlsx(".\\Input\\ARS_CMreg_Automation_input.xlsx",sheet = 'Sheet2')
#i<-length(Output_fields$Plan[!is.na(Output_fields$Plan)])

plans <- Output_fields$Plan

'%nin%' <- Negate(`%in%`)
#names(file_split)
for(i in 1: length(Output_fields$Plan)) {
  
  file_split <- Output_fields %>% filter(Plan %in% c(i))
  lender_split <- file_split$lender
  lender_split_1 <- file_split$Lender_Split_File
  
  column_names <- fields_name[,i] 
  
  
  if(Lender_Split_File =='Yes'){
    Arsdump <- Arsdump %>% filter(lender %in% c(lender_split))
  } else{
    
    product_split_1 <- lender_split$Product_Split_File
  }
    if(Product_Split_File=='Yes'){
      Arsdump <- Arsdump %>% filter(!is.na(Product))
    } else{
      
      Arsdump <- Arsdump
    }
    
  Arsdump_1 <- Arsdump %>% select(column_names)
  
  }


Input_file <- read.xlsx(".\\Input\\ARS_CMreg_Automation_input.xlsx",sheet = 'Sheet2')
  
  #base_1 <- lender_split$base
  
  allocation_vintage_1 <- lapply(strsplit(as.character(lender_split$allocation_vintage),","),as.character)
  
  allocation_vintage_1 <- unlist(allocation_vintage_1)
  
  mo_allocation_vintage_1 <- lender_split$mo_allocation_vintage
  