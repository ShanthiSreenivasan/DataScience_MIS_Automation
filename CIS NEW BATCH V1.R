rm(list = ls()) # Clear environment
Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
Sys.setenv('JAVA_HOME'="C:\\Program Files\\Java\\jdk1.8.0_172/")
library(data.table)
library(stringr)
library(openxlsx)
library(dplyr)
library(readxl)
library(excel.link) #Read xl with password
library(lubridate)
library(xtable)
library(mailR)
library(XLConnect)


setwd(Sys.getenv('CIS NEW LENDER BATCH'))

today <-Sys.Date()

source('.\\Input\\Lender Details\\Backup R data\\Function file.R')

hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")

thisDate = Sys.Date()
if(!dir.exists(paste("./Output/",thisDate,sep="")))
{
  dir.create(paste("./Output/",thisDate,sep=""))
} 


if(!dir.exists(paste("./Output/",thisDate,"/Batch to lender New cases",sep="")))
{
  dir.create(paste("./Output/",thisDate,"/Batch to lender New cases",sep=""))
} 


if(!dir.exists(paste("./Output/",thisDate,"/Payment file",sep="")))
{
  dir.create(paste("./Output/",thisDate,"/Payment file",sep=""))
} 

if(!dir.exists(paste("./Output/",thisDate,"/Flows file",sep="")))
{
  dir.create(paste("./Output/",thisDate,"/Flows file",sep=""))
} 

if(!dir.exists(paste("./Output/",thisDate,"/Others",sep="")))
{
  dir.create(paste("./Output/",thisDate,"/Others",sep=""))
} 


dir.name <- paste('.\\Output\\',thisDate,sep = '')

year_start <- paste("01-01-",format(Sys.Date(),"%Y")," 00:00:00",sep = '')

year_start <- dmy_hms(year_start)

# reading input dump file
dfCISdump=fread('./Input/cis_dump_new_cases.csv')

dfCISdump$CISserialno <- as.numeric(dfCISdump$CISserialno)

dfCISdump$subscription_date <- dmy_hm(dfCISdump$subscription_date)

wh_dump <- fread('./Input/CIS Dump_WH_Dump.csv')

wh_dump$lead_id <- as.numeric(wh_dump$lead_id)

wh_dump$subscription_date <- dmy_hm(wh_dump$subscription_date)

dflender =read.xlsx("./Input/Lender Details/Lender Details file.xlsx",sheet = 1)

dflendernames =read.xlsx("./Input/Lender Details/Lender Details file.xlsx",sheet = 2)

Add_cases=read.xlsx('./Input/Lender wise add cases.xlsx')

Add_cases$Lead_id <- as.numeric(Add_cases$Lead_id)

if (nrow(Add_cases)>0) {
  
  Add_cases <- left_join(Add_cases,wh_dump[,c("subscription_date","first_name",
                                              "last_name","lead_id","Account_No","product_status",
                                              "lender","product","account_status")],
                         by=c("Lead_id" = "lead_id"))
  
  
  Add_cases$S.No=""
  Add_cases$`Customer Name`=paste(Add_cases$first_name," ",Add_cases$last_name,sep = '')
  Add_cases$`Amount due for Bereau Clearance`=""
  Add_cases$`Allocation Status`=""
  Add_cases <- left_join(Add_cases,dflendernames,by=c("lender" = "lender_wh"))
  
  Add_cases <- Add_cases %>% select("S.No","subscription_date","Customer Name","Lead_id","Account_No",
                                    "product_status","lender_name_New","product","account_status","Amount due for Bereau Clearance",
                                    "Allocation Status")
  
 names(Add_cases) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")
  
}





dfunique_lenders=  unique(as.character(dflender$lender_name))

# file creation---------------------------------------------------------------
for (i in dfunique_lenders)
  
{
  print(i)
  # browser()
  #sheet1
  dfdata = dfCISdump %>% 
    filter(lender_name==i) %>% 
    filter (LDB=="Yes") %>% 
    filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))
  
  if (nrow(dfdata)>0) {
  dfdata$S.NO=""
  dfdata$Allocation_status=""
  dfdata$`Amount due for Bereau Clearance`=""
  dfdata$Customer_Name=paste(dfdata$first_name," ",dfdata$last_name,sep = '')
  
  dfdata1 = dfdata %>% select("S.NO",
                                    "subscription_date",
                                    "Customer_Name",
                                    "CISserialno",
                                    "Account_No",
                                    "prodstatus",
                                    "lender_name",
                                    "product_family_short",
                                    "account_status",
                                    "Amount due for Bereau Clearance", 
                                    "Allocation_status")

  
  
  names(dfdata1) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")
  
  Add_cases_batch <- Add_cases %>% filter(`Lender Name`==i)
  
  dfdata_New <- dfdata1 %>% filter(`CM Subscription Date` > year_start)
  dfdata_LTD <- dfdata1 %>% filter(`CM Subscription Date` < year_start)
  
  dfdata_New <- rbind(dfdata_New,Add_cases_batch) %>% distinct(`CM Lead ID`, .keep_all = T)
  
  
  if(nrow(dfdata_New)>0 & nrow(dfdata_LTD)>0){ 
    
    dfdata_New$S.No=1:nrow(dfdata_New)
    dfdata_LTD$S.No=1:nrow(dfdata_LTD)
    
    row_set <- dflender %>% filter(lender_name==i)
    Filename1 <-row_set$File_name
    
    Filename <- paste(dir.name,'\\Batch to lender New cases\\',Filename1,sep = '')
    FileCreate(dataset=dfdata_New,sheet_name="Priority Customers",Filename,
               sheet_name2 = "LTD STock",dataset2 = dfdata_LTD)
    
  } else if (nrow(dfdata_New) > 0 & nrow(dfdata_LTD) == 0) {
    
    dfdata_New$S.No=1:nrow(dfdata_New)
    
    row_set <- dflender %>% filter(lender_name==i)
    Filename1 <-row_set$File_name
    
    Filename <- paste(dir.name,'\\Batch to lender New cases\\',Filename1,sep = '')
    FileCreate(dataset=dfdata_New,sheet_name="Priority Customers",Filename)
    
  } else if (nrow(dfdata_New) == 0 & nrow(dfdata_LTD)>0) {
    
    dfdata_LTD$S.No=1:nrow(dfdata_LTD)
    
    row_set <- dflender %>% filter(lender_name==i)
    Filename1 <-row_set$File_name
    
    Filename <- paste(dir.name,'\\Batch to lender New cases\\',Filename1,sep = '')
    FileCreate(dataset=dfdata_LTD,sheet_name="LTD STock",Filename)
    
    
  } 
  
  } else {
    row_set <- dflender %>% filter(lender_name==i)
    Filename1 <-row_set$File_name1
    
    Filename <- paste(dir.name,'\\Others\\',Filename1,sep = '')
    FileCreate(dataset=dfdata,sheet_name="No leads",Filename)
    
}

}


#HDFC-----------------------------------------------------------------------------------------


dfhdfc= dfCISdump %>% 
  filter(lender_name=="HDFC") %>% 
  filter (LDB=="Yes") %>% 
  filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

dfhdfc$S.NO=""
dfhdfc$Allocation_status=""
dfhdfc$`Amount due for Bereau Clearance`=""
dfhdfc$Customer_Name=paste(dfhdfc$first_name," ",dfhdfc$last_name,sep = '')

dfhdfc1 = dfhdfc %>% select("S.NO",
                            "subscription_date",
                            "Customer_Name",
                            "CISserialno",
                            "Account_No",
                            "prodstatus",
                            "lender_name",
                            "product_family_short",
                            "account_status",
                            "Amount due for Bereau Clearance", 
                            "Allocation_status")



names(dfhdfc1) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")

dfhdfc_New <- dfhdfc1 %>% filter(`CM Subscription Date` > year_start)
dfhdfc_LTD <- dfhdfc1 %>% filter(`CM Subscription Date` < year_start)


# HDFC CC ----------------------------------------------------------------

dfhdfc_New_CC <- dfhdfc_New %>% filter(product %in% c("CC","SCC")) 
dfhdfc_LTD_CC <- dfhdfc_LTD %>% filter(product %in% c("CC","SCC")) 

HFDC_ADD_CASES_CC <- Add_cases %>% filter(`Lender Name`=="HDFC Bank")  %>% filter(product %in% c("CC","SCC")) 

dfhdfc_New_CC <- rbind(dfhdfc_New_CC,HFDC_ADD_CASES_CC) %>% distinct(`CM Lead ID`, .keep_all = T)


if(nrow(dfhdfc_New_CC)>0 & nrow(dfhdfc_LTD_CC)>0){ 
  
  dfhdfc_New_CC$S.No=1:nrow(dfhdfc_New_CC)
  dfhdfc_LTD_CC$S.No=1:nrow(dfhdfc_LTD_CC)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CC New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CC,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dfhdfc_LTD_CC)
  
} else if (nrow(dfhdfc_New_CC) > 0 & nrow(dfhdfc_LTD_CC) == 0) {
  
  dfhdfc_New_CC$S.No=1:nrow(dfhdfc_New_CC)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CC New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CC,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dfhdfc_New_CC) == 0 & nrow(dfhdfc_LTD_CC)>0) {
  
  dfhdfc_LTD_CC$S.No=1:nrow(dfhdfc_LTD_CC)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CC New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_LTD_CC,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dfhdfc_New_CC) == 0 & nrow(dfhdfc_LTD_CC) == 0) {
  

  Filename <- paste(dir.name,'\\Others\\HDFC Bank CC No leads found.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CC,sheet_name="No leads",Filename)
  
}



#HDFC CD ---------------------------------------------------------------------------

dfhdfc_New_CL <- dfhdfc_New %>% filter(product == "CL") 
dfhdfc_LTD_CL <- dfhdfc_LTD %>% filter(product == "CL") 

HFDC_ADD_CASES_CL <- Add_cases %>% filter(`Lender Name`=="HDFC Bank")  %>% filter(product == "CL") 

dfhdfc_New_CL <- rbind(dfhdfc_New_CL,HFDC_ADD_CASES_CL) %>% distinct(`CM Lead ID`, .keep_all = T)


if(nrow(dfhdfc_New_CL)>0 & nrow(dfhdfc_LTD_CL)>0){ 
  
  dfhdfc_New_CL$S.No=1:nrow(dfhdfc_New_CL)
  dfhdfc_LTD_CL$S.No=1:nrow(dfhdfc_LTD_CL)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CD New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CL,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dfhdfc_LTD_CL)
  
} else if (nrow(dfhdfc_New_CL) > 0 & nrow(dfhdfc_LTD_CL) == 0) {
  
  dfhdfc_New_CL$S.No=1:nrow(dfhdfc_New_CL)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CD New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CL,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dfhdfc_New_CL) == 0 & nrow(dfhdfc_LTD_CL)>0) {
  
  dfhdfc_LTD_CL$S.No=1:nrow(dfhdfc_LTD_CL)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank CL New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_LTD_CL,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dfhdfc_New_CL) == 0 & nrow(dfhdfc_LTD_CL) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\HDFC Bank CD No leads found.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_CL,sheet_name="No leads",Filename)
  
}


# HDFC RETAIL ----------------------------------------------------------------

dfhdfc_New_RETAIL <- dfhdfc_New %>% filter(!product %in% c("CC","SCC","CL")) 
dfhdfc_LTD_RETAIL <- dfhdfc_LTD %>% filter(!product %in% c("CC","SCC","CL")) 


HFDC_ADD_CASES_RETAIL <- Add_cases %>% filter(`Lender Name`=="HDFC")  %>% filter(!product %in% c("CC","SCC","CL")) 

dfhdfc_New_RETAIL <- rbind(dfhdfc_New_RETAIL,HFDC_ADD_CASES_RETAIL) %>% distinct(`CM Lead ID`, .keep_all = T)


if(nrow(dfhdfc_New_RETAIL)>0 & nrow(dfhdfc_LTD_RETAIL)>0){ 
  
  dfhdfc_New_RETAIL$S.No=1:nrow(dfhdfc_New_RETAIL)
  dfhdfc_LTD_RETAIL$S.No=1:nrow(dfhdfc_LTD_RETAIL)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank RETAIL New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_RETAIL,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dfhdfc_LTD_RETAIL)
  
} else if (nrow(dfhdfc_New_RETAIL) > 0 & nrow(dfhdfc_LTD_RETAIL) == 0) {
  
  dfhdfc_New_RETAIL$S.No=1:nrow(dfhdfc_New_RETAIL)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank RETAIL New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_New_RETAIL,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dfhdfc_New_RETAIL) == 0 & nrow(dfhdfc_LTD_RETAIL)>0) {
  
  dfhdfc_LTD_RETAIL$S.No=1:nrow(dfhdfc_LTD_RETAIL)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\HDFC Bank RETAIL New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_LTD_RETAIL,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dfhdfc_New_RETAIL) == 0 & nrow(dfhdfc_LTD_RETAIL) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\HDFC Bank RETAIL No leads found.xlsx',sep = '')
  FileCreate(dataset=dfhdfc_LTD_RETAIL,sheet_name="No leads",Filename)
  
}




#fullerton-----------------------------------------------------------------------------------------


dffull= dfCISdump %>% 
  filter(lender_name=="FULLERTON") %>% 
  filter (LDB=="Yes") %>% 
  filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

dffull$S.NO=""
dffull$Allocation_status=""
dffull$`Amount due for Bereau Clearance`=""
dffull$Customer_Name=paste(dffull$first_name," ",dffull$last_name,sep = '')

dffull1 = dffull %>% select("S.NO",
                            "subscription_date",
                            "Customer_Name",
                            "CISserialno",
                            "Account_No",
                            "prodstatus",
                            "lender_name",
                            "product_family_short",
                            "account_status",
                            "Amount due for Bereau Clearance", 
                            "Allocation_status")


names(dffull1) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")

dffull_New1 <- dffull1 %>% filter(`CM Subscription Date` > year_start)
dffull_LTD1<- dffull1 %>% filter(`CM Subscription Date` < year_start)

Full_Add_cases <- Add_cases %>% filter(`Lender Name`=="FULLERTON")

dffull_New1 <-rbind(dffull_New1,Full_Add_cases) %>% distinct(`CM Lead ID`, .keep_all = T)


dffull1_new_1 <- dffull_New1 %>% filter(str_detect(`Card/Loan No`, '[A-Za-z]'))
dffull_New <- anti_join(dffull_New1,dffull1_new_1,by=("Card/Loan No"))


dffull1_LTD_1 <- dffull_LTD1 %>% filter(str_detect(`Card/Loan No`, '[A-Za-z]'))
dffull_LTD <- anti_join(dffull_LTD1,dffull1_LTD_1,by=("Card/Loan No"))


if(nrow(dffull_New)>0 & nrow(dffull_LTD)>0){ 
  
  dffull_New$S.No=1:nrow(dffull_New)
  dffull_LTD$S.No=1:nrow(dffull_LTD)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Fullerton New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dffull_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dffull_LTD)
  
} else if (nrow(dffull_New) > 0 & nrow(dffull_LTD) == 0) {
  
  dffull_New$S.No=1:nrow(dffull_New)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Fullerton New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dffull_New,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dffull_New) == 0 & nrow(dffull_LTD)>0) {
  
  dffull_LTD$S.No=1:nrow(dffull_LTD)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Fullerton New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dffull_LTD,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dffull_New) == 0 & nrow(dffull_LTD) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\Fullerton No leads found.xlsx',sep = '')
  FileCreate(dataset=dffull_New,sheet_name="No leads",Filename)
  
}


#Kotak and pheonix ----------------------------------------------------------------


dfkotak= dfCISdump %>% 
  filter(lender_name %in% c("KOTAK","Phoenix")) %>% 
  filter (LDB=="Yes") %>% 
  filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

dfkotak$S.NO=""
dfkotak$Allocation_status=""
dfkotak$`Amount due for Bereau Clearance`=""
dfkotak$Customer_Name=paste(dfkotak$first_name," ",dfkotak$last_name,sep = '')

dfkotak1 = dfkotak %>% select("S.NO",
                            "subscription_date",
                            "Customer_Name",
                            "CISserialno",
                            "Account_No",
                            "prodstatus",
                            "lender_name",
                            "product_family_short",
                            "account_status",
                            "Amount due for Bereau Clearance", 
                            "Allocation_status")

names(dfkotak1) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")

dfkotak_New <- dfkotak1 %>% filter(`CM Subscription Date` > year_start)
dfkotak_LTD <- dfkotak1 %>% filter(`CM Subscription Date` < year_start)

Kotak_Add_cases <- Add_cases %>% filter(`Lender Name` %in% c("KOTAK","Phoenix"))

dfkotak_New <- rbind(dfkotak_New,Kotak_Add_cases) %>% distinct(`CM Lead ID`, .keep_all = T)


# Files 

Kotak_details <- read_excel('.\\Input\\Lender Details\\kotak & phoenix.xlsx',sheet = 'Sheet1')

P_K_New <- dfkotak_New %>% mutate(Lender_code = str_replace(substr(dfkotak_New$`Card/Loan No`,1,5),"` ",""))
P_K_LTD <- dfkotak_LTD %>% mutate(Lender_code = str_replace(substr(dfkotak_LTD$`Card/Loan No`,1,5),"` ","")) 


P_K_New_final <- left_join(P_K_New,
                                 Kotak_details[,c("CONTRACT NO","FPR NAME")],
                                 by = c('Lender_code'='CONTRACT NO')) %>% 
  distinct(`CM Lead ID`, .keep_all = T) 


P_K_LTD_final <- left_join(P_K_LTD,
                             Kotak_details[,c("CONTRACT NO","FPR NAME")],
                             by = c('Lender_code'='CONTRACT NO')) %>% 
  distinct(`CM Lead ID`, .keep_all = T) 


Ganesh_New <- P_K_New_final %>% 
  filter(grepl('Ganesh',`FPR NAME`)) %>% 
  select(-Lender_code,
         -`FPR NAME`)

Ganesh_LTD <- P_K_LTD_final %>% 
  filter(grepl('Ganesh',`FPR NAME`)) %>% 
  select(-Lender_code,
         -`FPR NAME`)

jacob_New <- P_K_New_final %>% 
  filter(grepl('JACOB',`FPR NAME`))%>% 
  select(-Lender_code,
         -`FPR NAME`)

jacob_LTD <- P_K_LTD_final %>% 
  filter(grepl('JACOB',`FPR NAME`))%>% 
  select(-Lender_code,
         -`FPR NAME`)


jacob_New_1 <-rbind(jacob_New,dffull1_new_1)

jacob_LTD_1 <-rbind(jacob_LTD,dffull1_LTD_1)


# Ganesh File 

if(nrow(Ganesh_New)>0 & nrow(Ganesh_LTD)>0){ 
  
  Ganesh_New$S.No=1:nrow(Ganesh_New)
  Ganesh_LTD$S.No=1:nrow(Ganesh_LTD)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix & Kotak New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=Ganesh_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = Ganesh_LTD)
  
} else if (nrow(Ganesh_New) > 0 & nrow(Ganesh_LTD) == 0) {
  
  Ganesh_New$S.No=1:nrow(Ganesh_New)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix & Kotak New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=Ganesh_New,sheet_name="Priority Customers",Filename)
  
} else if (nrow(Ganesh_New) == 0 & nrow(Ganesh_LTD)>0) {
  
  Ganesh_LTD$S.No=1:nrow(Ganesh_LTD)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix & Kotak New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=Ganesh_LTD,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(Ganesh_New) == 0 & nrow(Ganesh_LTD) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\Phoenix & Kotak No leads found.xlsx',sep = '')
  FileCreate(dataset=Ganesh_New,sheet_name="No leads",Filename)
  
}


#jocab file 


if(nrow(jacob_New_1)>0 & nrow(jacob_LTD_1)>0){ 
  
  jacob_New_1$S.No=1:nrow(jacob_New_1)
  jacob_LTD_1$S.No=1:nrow(jacob_LTD_1)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=jacob_New_1,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = jacob_LTD_1)
  
} else if (nrow(jacob_New_1) > 0 & nrow(jacob_LTD_1) == 0) {
  
  jacob_New_1$S.No=1:nrow(jacob_New_1)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=jacob_New_1,sheet_name="Priority Customers",Filename)
  
} else if (nrow(jacob_New_1) == 0 & nrow(jacob_LTD_1)>0) {
  
  jacob_LTD_1$S.No=1:nrow(jacob_LTD_1)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\Phoenix New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=jacob_LTD_1,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(jacob_New_1) == 0 & nrow(jacob_LTD_1) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\Phoenix No leads found.xlsx',sep = '')
  FileCreate(dataset=jacob_New_1,sheet_name="No leads",Filename)
  
}


#SBI NON LDB CASES---------------------------------------------------------

dfsbi= dfCISdump %>% 
  filter(lender_name=="SBI Cards") %>% 
 # filter (LDB=="No") %>% 
  filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

dfsbi$S.NO=""
dfsbi$Allocation_status=""
dfsbi$`Amount due for Bereau Clearance`=""
dfsbi$Customer_Name=paste(dfsbi$first_name," ",dfsbi$last_name,sep = '')

dfsbi = dfsbi %>% select("S.NO",
                            "subscription_date",
                            "Customer_Name",
                            "CISserialno",
                            "Account_No",
                            "prodstatus",
                            "lender_name",
                            "product_family_short",
                            "account_status",
                            "Amount due for Bereau Clearance", 
                            "Allocation_status")


names(dfsbi) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")


dfsbi_New <- dfsbi %>% filter(`CM Subscription Date` > year_start)
dfsbi_LTD <- dfsbi %>% filter(`CM Subscription Date` < year_start)


dfsbi_ADD_CASES_CC <- Add_cases %>% filter(`Lender Name`=="SBI Cards")  

dfsbi_New <- rbind(dfsbi_New,dfsbi_ADD_CASES_CC) %>% distinct(`CM Lead ID`, .keep_all = T)


if(nrow(dfsbi_New)>0 & nrow(dfsbi_LTD)>0){ 
  
  dfsbi_New$S.No=1:nrow(dfsbi_New)
  dfsbi_New$S.No=1:nrow(dfsbi_New)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfsbi_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dfsbi_LTD)
  
} else if (nrow(dfsbi_New) > 0 & nrow(dfsbi_LTD) == 0) {
  
  dfsbi_New$S.No=1:nrow(dfsbi_New)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfsbi_New,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dfsbi_New) == 0 & nrow(dfsbi_LTD)>0) {
  
  dfsbi_LTD$S.No=1:nrow(dfsbi_LTD)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfsbi_LTD,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dfsbi_New) == 0 & nrow(dfsbi_LTD) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\SBI Cards No leads found.xlsx',sep = '')
  FileCreate(dataset=dfsbi_New,sheet_name="No leads",Filename)
  
}



#SBI Payments file-----------------------------------------------------------

dfpayment_SBI <- fread(paste('./Input/Payments file/SBI CC Payments - ',today,'.csv',sep = ''))
SBI_file <-paste("Output\\",today,"\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx",sep='')

Account <- read_excel(paste('./Input/Payments file/Payments encription.xlsx',sep = ''))

df <- left_join(dfpayment_SBI, Account, by = c('ACCOUNT NUMBER'='account_no'))

df <- df %>% select(c("CUSTOMER NAME","Account_No","ACCOUNT TYPE","AGENCY NAME",
                      "INSTALLMENT","GROSS PAYMENT","NET PAYMENT","PAYMENT MODE",
                      "PAYMENT DATE","TRANSACTION DETAILS","EMAIL ID","MOBILE NO","COLLECTOR NAME",
                      "REGION"))
names(df) <- c("CUSTOMER NAME","ACCOUNT NUMBER","ACCOUNT TYPE","AGENCY NAME",
               "INSTALLMENT","GROSS PAYMENT","NET PAYMENT","PAYMENT MODE",
               "PAYMENT DATE","TRANSACTION DETAILS","EMAIL ID","MOBILE NO","COLLECTOR NAME",
               "REGION")


df <- left_join(df, Account, by = c('TRANSACTION DETAILS'='account_no'))

df <- df %>% select(c("CUSTOMER NAME","ACCOUNT NUMBER","ACCOUNT TYPE","AGENCY NAME",
                      "INSTALLMENT","GROSS PAYMENT","NET PAYMENT","PAYMENT MODE",
                      "PAYMENT DATE","Account_No","EMAIL ID","MOBILE NO","COLLECTOR NAME",
                      "REGION"))
names(df) <- c("CUSTOMER NAME","ACCOUNT NUMBER","ACCOUNT TYPE","AGENCY NAME",
               "INSTALLMENT","GROSS PAYMENT","NET PAYMENT","PAYMENT MODE",
               "PAYMENT DATE","TRANSACTION DETAILS","EMAIL ID","MOBILE NO","COLLECTOR NAME",
               "REGION")


dfpayment_SBI <- df

dfpayment_SBI$`TRANSACTION DETAILS`[dfpayment_SBI$`TRANSACTION DETAILS`==""] <- '` NA'
dfpayment_SBI$`TRANSACTION DETAILS`[is.na(dfpayment_SBI$`TRANSACTION DETAILS`)] <- '` NA'

dfpayment_SBI$`PAYMENT MODE`[dfpayment_SBI$`PAYMENT MODE`==""] <- 'Online PG'
dfpayment_SBI$`PAYMENT MODE`[is.na(dfpayment_SBI$`PAYMENT MODE`)] <- 'Online PG'

myWorkbook <- XLConnect::loadWorkbook(SBI_file)
numberofsheets <- length(getSheets(myWorkbook))
sheet <-excel_sheets(SBI_file)

if (numberofsheets > 1) {
  C_SBI_New <- read_excel(SBI_file,sheet = "Priority Customers")
  C_SBI_LTD <- read_excel(SBI_file,sheet = "LTD STock")
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=C_SBI_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",sheet_name3="Payments",dataset2 = C_SBI_LTD,dataset3=dfpayment_SBI)
} else {
    
  C_SBI_New <- read_excel(SBI_file,sheet = "Priority Customers")
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=C_SBI_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "Payments",dataset2 = dfpayment_SBI)
  
  
  }


#SCB Payment file--------------------------------------------------------------------------------

scb_payment_file <- fread(paste('./Input/Payments file/SCB Payments - ',today,'.csv',sep = ''))

  
df_scb <- left_join(scb_payment_file, Account, by = c('Account No'='account_no'))


df_scb <- df_scb %>% select("id","lead_id","Customer Name","Account_No","lender","product",       
                             "account_status","product_status","amount_paid","payment_date")


names(df_scb)<-c("id","lead_id","Customer Name","Account No","lender","product","account_status","product_status","amount_paid","payment_date")

colnames(wh_dump) <- make.unique(names(wh_dump))

scb_ptp <- wh_dump %>%
  filter(lender %in% c('SCB') & is_ldb %in% c(1) & latest_dispo %in% c('PTP')) %>% 
  select(lead_id, Account_No, product) %>% 
  mutate(feedback = 'PTP')


if (nrow(df_scb)>0) {
 
  Filename <- paste(dir.name,'\\Payment file\\SCB Payment File.xlsx',sep = '')
  FileCreate(dataset=df_scb,sheet_name="SCB Payment File",Filename,
             sheet_name2 = "Feedback",dataset2 = scb_ptp) 
} else {
  
  Filename <- paste(dir.name,'\\Others\\SCB Payment File.xlsx',sep = '')
  FileCreate(dataset=df_scb,sheet_name="SCB Payment File",Filename,
             sheet_name2 = "Feedback",dataset2 = scb_ptp)
}



#HDFC Payments------------------------------------------------------------------

hdfc_payments <- fread(paste('./Input/Payments file/HDFC Payments - ',today,'.csv',sep=''))

df_hdfc <- left_join(hdfc_payments, Account, by = c('Loan No'='account_no'))

df_hdfc <- df_hdfc %>% select("lead_id","Account_No","Customer Name","Product","Paid Amount","TRANSACTION / NARRATION DETAIL",
                              "Mode of Payment","DUMMY A/C NO","Payment Date","Location","Account Status")


names(df_hdfc) <- c("lead_id","Loan No","Customer Name","Product","Paid Amount","TRANSACTION / NARRATION DETAIL",
                    "Mode of Payment","DUMMY A/C NO","Payment Date","Location","Account Status")


df_hdfc <- left_join(df_hdfc, Account, by = c('TRANSACTION / NARRATION DETAIL'='account_no'))


df_hdfc <- df_hdfc %>% select("lead_id","Loan No","Customer Name","Product","Paid Amount","Account_No",
                              "Mode of Payment","DUMMY A/C NO","Payment Date","Location","Account Status")

names(df_hdfc) <- c("lead_id","Loan No","Customer Name","Product","Paid Amount","TRANSACTION / NARRATION DETAIL",
                    "Mode of Payment","DUMMY A/C NO","Payment Date","Location","Account Status")



df_hdfc <- left_join(df_hdfc, Account, by = c('DUMMY A/C NO'='account_no'))

df_hdfc <- df_hdfc %>% select("lead_id","Loan No","Customer Name","Product","Paid Amount","TRANSACTION / NARRATION DETAIL",
                              "Mode of Payment","Account_No","Payment Date","Location","Account Status")


names(df_hdfc) <- c("lead_id","Loan No","Customer Name","Product","Paid Amount","TRANSACTION / NARRATION DETAIL",
                    "Mode of Payment","DUMMY A/C NO","Payment Date","Location","Account Status")


df_hdfc$`TRANSACTION / NARRATION DETAIL`[df_hdfc$`TRANSACTION / NARRATION DETAIL`==""] <- '` NA'
df_hdfc$`TRANSACTION / NARRATION DETAIL`[is.na(df_hdfc$`TRANSACTION / NARRATION DETAIL`)] <- '` NA'
df_hdfc$`DUMMY A/C NO`[df_hdfc$`DUMMY A/C NO`==""] <- '` NA'
df_hdfc$`DUMMY A/C NO`[is.na(df_hdfc$`DUMMY A/C NO`)] <- '` NA'
df_hdfc$`Mode of Payment`[df_hdfc$`Mode of Payment`==""] <- 'Online PG'
df_hdfc$`Mode of Payment`[is.na(df_hdfc$`Mode of Payment`)] <- 'Online PG'


hdfc_payments <- df_hdfc

#hdfc_payments$`Paid Amount` <- as.character(hdfc_payments$`Paid Amount`)

#Filter CL Cases
hdfc_payments_cl <- hdfc_payments[hdfc_payments$Product == 'CL',]
if(nrow(hdfc_payments_cl)>0){
  
  hdfc_payments_cl_path <- paste(dir.name,'\\Payment file\\HDFC Retail Payment File with CL.xlsx',sep = '')
  FileCreate(dataset=hdfc_payments_cl,sheet_name="Payments",hdfc_payments_cl_path)
  
} else {
  
  hdfc_payments_cl_path <- paste(dir.name,'\\Others\\HDFC Retail Payment File with CL.xlsx',sep = '')
  FileCreate(dataset=hdfc_payments_cl,sheet_name="Payments",hdfc_payments_cl_path)
  
}

#Remove CL from HDFC Payments
hdfc_payments <- hdfc_payments[hdfc_payments$Product != 'CL',]
#Write file

if (nrow(hdfc_payments)>0) {

  hdfc_payments_path <- paste(dir.name,'\\Payment file\\HDFC Retail Payment File.xlsx',sep = '')
  FileCreate(dataset=hdfc_payments,sheet_name="Payments",hdfc_payments_path)
  
}else{
  
  hdfc_payments_path <- paste(dir.name,'\\Others\\HDFC Retail Payment File.xlsx',sep = '')
  FileCreate(dataset=hdfc_payments,sheet_name="Payments",hdfc_payments_path)
}

#Kotak and phoniex Payment file------------------------------------------------------

Kotak_details <- read_excel('.\\Input\\Lender Details\\kotak & phoenix.xlsx',sheet = 'Sheet1')

Kotak_Path <- paste('.\\Input\\Payments file\\KOTAK & Phoenix Payments - ',today,'.csv',sep = "")

Kotak_Payment <-fread(Kotak_Path)



df_Kotak <- left_join(Kotak_Payment, Account, by = c('Account Number'='account_no'))


df_Kotak <- df_Kotak %>% select("lead_id","Name","Account_No","lender","product","account_status",
                                "amount_paid","payment_date","mode")


names(df_Kotak)<-c("lead_id","Name","Account Number","lender","product","account_status",
                   "amount_paid","payment_date","mode")

df_Kotak$mode[df_Kotak$mode==""] <- 'Online PG'
df_Kotak$mode[is.na(df_Kotak$mode)] <- 'Online PG'

Kotak_Payment <- df_Kotak


Kotak_Payment$lead_id <- as.numeric(Kotak_Payment$lead_id)

wh_dump$lead_id <- as.numeric(wh_dump$lead_id)

Kotak_Payment1 <- left_join(
  Kotak_Payment,
  wh_dump[,c('lead_id','Account_No')],
  by = 'lead_id')

Kotak_Payment1 <-Kotak_Payment1 %>%   mutate(Lender_code = str_replace(substr(Kotak_Payment1$Account_No,1,5),"` ",""))

P_K_final <- left_join(Kotak_Payment1,
                       Kotak_details[,c("CONTRACT NO","FPR NAME")],
                       by = c('Lender_code'='CONTRACT NO')) 


Ganesh <- P_K_final %>% 
  filter(grepl('Ganesh',`FPR NAME`)) %>% 
  select(-Lender_code,
         -`FPR NAME`) %>% select("lead_id","Name","Account_No","lender","product","account_status","amount_paid","payment_date","mode")

jacob_1 <- P_K_final %>% 
  filter(grepl('JACOB',`FPR NAME`))%>% 
  select(-Lender_code,
         -`FPR NAME`) %>% select("lead_id","Name","Account_No","lender","product","account_status","amount_paid","payment_date","mode")

if (nrow(Ganesh)>0) {
  Filename <- paste(dir.name,'\\Payment file\\Kotak and Phoneix Payment File.xlsx',sep = '')
  FileCreate(dataset=Ganesh,sheet_name="Kotak & phoneix",Filename)
} else  {
  
  Filename <- paste(dir.name,'\\Others\\Kotak and Phoneix Payment File.xlsx',sep = '')
  FileCreate(dataset=Ganesh,sheet_name="Kotak & phoneix",Filename)
}


if (nrow(jacob_1)>0) {
  Filename <- paste(dir.name,'\\Payment file\\Phoneix Payment file.xlsx',sep = '')
  FileCreate(dataset=jacob_1,sheet_name="phoneix",Filename)
} else {
  
  Filename <- paste(dir.name,'\\Others\\Phoneix Payment file.xlsx',sep = '')
  FileCreate(dataset=jacob_1,sheet_name="phoneix",Filename)
  
}
#L & T Finance-----------------------------------------------------------------------------------------


dfLNT= dfCISdump %>% 
  filter(lender_name=="L & T Finance") %>% 
  filter (LDB=="Yes") %>% 
  filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

dfLNT$S.NO=""
dfLNT$Allocation_status=""
dfLNT$`Amount due for Bereau Clearance`=""
dfLNT$Customer_Name=paste(dfLNT$first_name," ",dfLNT$last_name,sep = '')

dfLNT1 = dfLNT %>% select("S.NO",
                          "subscription_date",
                          "Customer_Name",
                          "CISserialno",
                          "Account_No",
                          "prodstatus",
                          "lender_name",
                          "product_family_short",
                          "account_status",
                          "Amount due for Bereau Clearance", 
                          "Allocation_status")


names(dfLNT1) <- c("S.No","CM Subscription Date","Customer Name","CM Lead ID","Card/Loan No","CM Prod status","Lender Name","product","Account Status","Amount due for Bereau Clearance in (Rs.)","Allocation Status")

dfLNT_New1 <- dfLNT1 %>% filter(`CM Subscription Date` > year_start)
dfLNT_LTD1<- dfLNT1 %>% filter(`CM Subscription Date` < year_start)

LNT_Add_cases <- Add_cases %>% filter(`Lender Name`=="L & T Finance")

dfLNT_New1 <-rbind(dfLNT_New1,LNT_Add_cases) %>% distinct(`CM Lead ID`, .keep_all = T)


dfLNT1_new_1 <- dfLNT_New1 %>% filter(str_detect(`Card/Loan No`, '[A-Za-z]'))
dfLNT_New <- anti_join(dfLNT_New1,dfLNT1_new_1,by=("Card/Loan No"))


dfLNT1_LTD_1 <- dfLNT_LTD1 %>% filter(str_detect(`Card/Loan No`, '[A-Za-z]'))
dfLNT_LTD <- anti_join(dfLNT_LTD1,dfLNT1_LTD_1,by=("Card/Loan No"))


if(nrow(dfLNT_New)>0 & nrow(dfLNT_LTD)>0){ 
  
  dfLNT_New$S.No=1:nrow(dfLNT_New)
  dfLNT_LTD$S.No=1:nrow(dfLNT_LTD)
  
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\L & T Finance New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfLNT_New,sheet_name="Priority Customers",Filename,
             sheet_name2 = "LTD STock",dataset2 = dfLNT_LTD)
  
} else if (nrow(dfLNT_New) > 0 & nrow(dfLNT_LTD) == 0) {
  
  dfLNT_New$S.No=1:nrow(dfLNT_New)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\L & T Finance New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfLNT_New,sheet_name="Priority Customers",Filename)
  
} else if (nrow(dfLNT_New) == 0 & nrow(dfLNT_LTD)>0) {
  
  dfLNT_LTD$S.No=1:nrow(dfLNT_LTD)
  
  Filename <- paste(dir.name,'\\Batch to lender New cases\\L & T Finance New cases & LTD cases.xlsx',sep = '')
  FileCreate(dataset=dfLNT_LTD,sheet_name="LTD STock",Filename)
  
  
} else if (nrow(dfLNT_New) == 0 & nrow(dfLNT_LTD) == 0) {
  
  
  Filename <- paste(dir.name,'\\Others\\L & T Finance No leads found.xlsx',sep = '')
  FileCreate(dataset=dfLNT_New,sheet_name="No leads",Filename)
  
}



#upolad one file----------------------------------------------------------------------------------------


dfup1 <- read.xlsx("./Input/Lender Details/Upload 1 Lender file.xlsx",sheet = 1)

dfdata <- list.files(paste(".\\Output\\",today,"\\Batch to lender New cases",sep=''), pattern=NULL, all.files=FALSE,
                     full.names=FALSE)


data1 <- c()

for(i in dfup1$lender_name){
  
  #browser()
  fname <- paste(".\\Output\\",today,"\\Batch to lender New cases\\",i,sep ='')
  
  if(file.exists(fname)){
  myWorkbook <- XLConnect::loadWorkbook(fname)
  numberofsheets <- length(getSheets(myWorkbook))
  
  
  
  if (numberofsheets>1) {
    
    x <- read_excel(fname,sheet = 1) %>%  select(`CM Lead ID`,`Lender Name`) %>% mutate(Type="Priority Customers")
    y <- read_excel(fname,sheet = 2) %>%  select(`CM Lead ID`,`Lender Name`) %>% mutate(Type="LTD STock")
    data1 <- rbind(data1,x,y)
    
  } else {
    
    x <- read_excel(fname,sheet = 1) %>%  select(`CM Lead ID`,`Lender Name`)%>% mutate(Type="Priority Customers")
    data1 <- rbind(data1,x)
    
  }
  



fname <- paste(".\\Output\\",today,"\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx",sep ='')
myWorkbook <- XLConnect::loadWorkbook(fname)
numberofsheets <- length(getSheets(myWorkbook))

if (numberofsheets>2) {
  
  x <- read_excel(fname,sheet = 1) %>%  select(`CM Lead ID`,`Lender Name`) %>% mutate(Type="Priority Customers")
  y <- read_excel(fname,sheet = 2) %>%  select(`CM Lead ID`,`Lender Name`) %>% mutate(Type="LTD STock")
  data1 <- rbind(data1,x,y)
  
} else if ((numberofsheets>1)) {
  
  x <- read_excel(fname,sheet = 1) %>%  select(`CM Lead ID`,`Lender Name`)%>% mutate(Type="Priority Customers")
  data1 <- rbind(data1,x)
  
}

 }
}

data1$`CM Lead ID` <- as.numeric(data1$`CM Lead ID`)
dfCISdump$CISserialno <- as.numeric(dfCISdump$CISserialno)

upload_one <- left_join(data1,dfCISdump[,c("CISserialno","subscription_date","LDB","prodstatus","Facilitystatus","product_family_short",
                                         "Account_No")],
                        by = c("CM Lead ID"="CISserialno"))

upload_one <- upload_one[ !is.na(upload_one$Account_No),]


upload_one <- upload_one %>% filter(upload_one$Facilitystatus == "00 To Start Action") 

upload_one <- upload_one %>% filter(upload_one$prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220"))

upload_one <- upload_one  %>% distinct(`CM Lead ID`, .keep_all = T) 

upload_path <- paste(dir.name,'\\Upload one.xlsx',sep = '')
FileCreate(dataset=upload_one,sheet_name="upload one",upload_path)

#<---------------Done------------------------>

