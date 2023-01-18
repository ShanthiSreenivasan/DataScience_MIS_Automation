#Remvove existing List
rm(list=ls())

#Set working directory
#setwd('Z:\\Amrish\\CIS')
Sys.setenv('JAVA_HOME'="C:\\Program Files\\Java\\jdk1.8.0_172/")
setwd(Sys.getenv('CIS NEW LENDER BATCH'))

#Load Library
library(stringr)
library(dplyr)
library(readxl)
library(excel.link) #Read xl with password
library(lubridate)
library(mailR)
library(xtable)
library(data.table)
library(htmlTable)
library(openxlsx)
library(XLConnect)


today <-Sys.Date()
source('.\\Input\\Lender Details\\Backup R data\\Function file.R')

Count_Mail <- read_excel('.\\Input\\MailMaster_Batch_New_zip.xlsx',sheet = 'Count & NLDB')

mail <- function(sender,recepients,myMessage,msg,cclist,filename=NULL){
  
  send.mail(from = sender,
            to = c(recepients),
            subject = myMessage,
            html = TRUE,
            body = str_interp(msg),
            cc = c(cclist),
            smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                        user.name = 'AKIAXJKXCWJGAXCHHIGY',
                        passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
            authenticate = TRUE,
            attach.files = filename,
            send = TRUE)
  
  
}


path<- paste(".\\Output\\",Sys.Date(),"\\Batch to lender New cases",sep ="")


New_batch_1 <- list.files(path,pattern = "xlsx$",
                        all.files = F,full.names = F)

New_batch <- New_batch_1[!grepl("SBI Cards New cases",New_batch_1)]

data1 <- c()

for(i in New_batch){
  
  #browser()
  path <- path<- paste(".\\Output\\",Sys.Date(),"\\Batch to lender New cases\\",i,sep ="")
  myWorkbook <- XLConnect::loadWorkbook(path)
  numberofsheets <- length(getSheets(myWorkbook))
  
  if (numberofsheets>1) {
    
    x <- read_excel(path,sheet = 1) 
    y <- read_excel(path,sheet = 2)
    z <- rbind(x,y)
    data <- nrow(z)
    x <- data.frame(File_name=i,count=data)
    
    data1 <- rbind(data1,x)
    
  } else if ((numberofsheets>0 & numberofsheets<2)) {
    
    x <- read_excel(path,sheet = 1)
    data <- nrow(x)
    x <- data.frame(File_name=i,count=data)
    data1 <- rbind(data1,x)
    
  }
}


path <- path<- paste(".\\Output\\",Sys.Date(),"\\Batch to lender New cases\\SBI Cards New cases & LTD cases.xlsx",sep ="")
myWorkbook <- XLConnect::loadWorkbook(path)
numberofsheets <- length(getSheets(myWorkbook))

if (numberofsheets>2) {
  
  x <- read_excel(path,sheet = 1)
  y <- read_excel(path,sheet = 2)
  z <- rbind(x,y)
  data <- nrow(z)
  x <- data.frame(File_name="SBI Cards New cases & LTD cases.xlsx",count=data)
  
} else if ((numberofsheets>1 & numberofsheets<3)) {
  
  x <- read_excel(path,sheet = 1)
  data <- nrow(x)
  x <- data.frame(File_name="SBI Cards New cases & LTD cases.xlsx",count=data)
  data1 <- rbind(data1,x)
  
} else if ((numberofsheets == 1)) {
  
  x <- read_excel(path,sheet = 1)
  data <- nrow(x)
  x <- data.frame(File_name="SBI Cards New cases & LTD cases.xlsx",count=data)
  data1 <- rbind(data1,x)
  
}



path<- paste(".\\Output\\",Sys.Date(),"\\Payment file",sep ="")

Payment_Batch <- list.files(path,pattern = "xlsx$",
                           all.files = F,full.names = F)



for(i in Payment_Batch){
  
  #browser()
  
  path <- path<- paste(".\\Output\\",Sys.Date(),"\\Payment file\\",i,sep ="")
  File_data <- read_excel(path)
  data <- nrow(File_data)
  x <- data.frame(File_name=i,count=data,stringsAsFactors = FALSE)
  
  data1 <- rbind(data1,x)
  
}


upload_path <- paste('.\\Output\\',Sys.Date(),'\\Count File.xlsx',sep = '')
FileCreate(dataset=data1,sheet_name="Count File",upload_path)


xx <- read_excel(upload_path,col_names = T)
df1 <- xx

df1 <- data.frame(df1)

df1$File_name <- str_replace(df1$File_name,".xlsx","")
df1$count <- as.numeric(df1$count)

df1 <- rbind(df1, c("Total", sum(df1$count)))

color_cells <- function(df, var){
  out <- ifelse(df[, var] == 0, 
                paste0('<div id="red">', df[, var], '</div>'),
                df[, var])
}

# apply coloring function to each column you want
df1$count <- color_cells(df = df1, var= "count")
df1$count <- as.character(df1$count)

Y <- xtable(df1)

message <- str_interp("BATCH COUNT ON - ${today}")

lapply(Count_Mail$`File Name`,function(z){
  #browser()
  # sender <- "referrals@creditmantri.com"
  sender <- 'Ops-cis@creditmantri.com'
  recepients <- str_split(Count_Mail$To[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  myMessage <- paste(Count_Mail$Subject[Count_Mail$`File Name` == z]," ",Sys.Date(),sep = '')
  cclist = str_split(Count_Mail$cc[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  msg <- paste('
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<style>
table {font-family:  Verdana, Geneva, sans-serif; font-size: 11px;
width = 100%; border: 1px solid black; border-collapse: collapse;
text-align: center; padding: 5px;}
th {height = 12px;background-color: #4CAF50;color: white;}
td {background-color: #FFF;}
#red{
  background-color:red;
}
</style>
</head>
<body>
  <h3>${message}</h3> 
  <p>${print(Y,type =\'html\', 
      hline.after=-1:3,
      sanitize.text.function = function(x) x)} </p>
  </body>
</html>')

  filename <- paste(".\\Output\\",today,"\\",z,sep = "")
  if(z == "Count File.xlsx"){
  mail(sender,recepients,myMessage,msg,cclist)
  }
})



# -------------------------------------------------------------------

Subcrib <-Sys.Date()-1
#Subcrib <-Subcrib," 00:00:00",sep = '')
sub <-Sys.Date()-2
today <- Sys.Date()

# reading input dump file
dfCISdump=read.csv("./Input/cis_dump_new_cases.csv",stringsAsFactors = FALSE) %>% filter(LDB == "Yes")
dfCISdump<-dfCISdump %>% filter(!lender_name %in% c("AU SMALL FINANCE BANK LIMITED","DEUTSCHE BANK","TATA Capital",
                                           "Reliance  Asset Reconstruction Company Limited","HOME CREDIT","AU FINANCIERS"))
dfCISdump$subscription_date <-dmy_hm(dfCISdump$subscription_date)
dfCISdump$subscription_date <-as.Date(dfCISdump$subscription_date)

Day_1_not_received <-dfCISdump %>% filter(subscription_date >= sub & subscription_date < Subcrib) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action") %>% mutate(Flag="Day-1")
Day_1_not_sent <-dfCISdump %>% filter(subscription_date >= sub & subscription_date < Subcrib) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action") %>% mutate(Flag="Day-1")


Day_1 <- Day_1_not_received %>% group_by(lender_name) %>% summarise("day-1"=n())

x <- format(Subcrib,"%b-%d")
names(Day_1) <- c("lender_name",x) 

Day_1_Not <- Day_1_not_sent %>% group_by(lender_name) %>% summarise("day-1"=n())

x <- format(Subcrib,"%b-%d")
names(Day_1_Not) <- c("lender_name",x) 




Day_2_not_received <-dfCISdump %>% filter(subscription_date >= sub-1 & subscription_date < sub) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-2")
Day_2_not_sent <-dfCISdump %>% filter(subscription_date >= sub-1 & subscription_date < sub) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-2")


Day_2 <- Day_2_not_received %>% group_by(lender_name) %>% summarise("day-2"=n())

x <- format(sub,"%b-%d")
names(Day_2) <- c("lender_name",x)

Day_2_Not <- Day_2_not_sent %>% group_by(lender_name) %>% summarise("day-2"=n())

x <- format(sub,"%b-%d")
names(Day_2_Not) <- c("lender_name",x)


Day_3_not_received <-dfCISdump %>% filter(subscription_date >= sub-2 & subscription_date < sub-1) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-3")
Day_3_not_sent <-dfCISdump %>% filter(subscription_date >= sub-2 & subscription_date < sub-1) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-3")


Day_3 <- Day_3_not_received %>% group_by(lender_name) %>% summarise("day-3"=n())

Day_3_Not <- Day_3_not_sent %>% group_by(lender_name) %>% summarise("day-3"=n())

x <- format(sub-1,"%b-%d")
names(Day_3) <- c("lender_name",x)
names(Day_3_Not) <- c("lender_name",x)


Day_4_not_received <-dfCISdump %>% filter(subscription_date >= sub-3 & subscription_date < sub-2) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-4")
Day_4_not_sent <-dfCISdump %>% filter(subscription_date >= sub-3 & subscription_date < sub-2) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-4")


Day_4 <- Day_4_not_received %>% group_by(lender_name) %>% summarise("day-4"=n())

Day_4_Not <- Day_4_not_sent %>% group_by(lender_name) %>% summarise("day-4"=n())

x <- format(sub-2,"%b-%d")
names(Day_4) <- c("lender_name",x)
names(Day_4_Not) <- c("lender_name",x)





Day_5_not_received <-dfCISdump %>% filter(subscription_date >= sub-4 & subscription_date < sub-3) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-5")
Day_5_not_sent <-dfCISdump %>% filter(subscription_date >= sub-4 & subscription_date < sub-3) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-5")


Day_5 <- Day_5_not_received %>% group_by(lender_name) %>% summarise("day-5"=n())

Day_5_Not <- Day_5_not_sent %>% group_by(lender_name) %>% summarise("day-5"=n())

x <- format(sub-3,"%b-%d")
names(Day_5) <- c("lender_name",x)
names(Day_5_Not) <- c("lender_name",x)



Day_6_not_received <-dfCISdump %>% filter(subscription_date >= sub-5 & subscription_date < sub-4) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-6")
Day_6_not_sent <-dfCISdump %>% filter(subscription_date >= sub-5 & subscription_date < sub-4) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-7")


Day_6 <- Day_6_not_received %>% group_by(lender_name) %>% summarise("day-6"=n())

Day_6_Not <- Day_6_not_sent %>% group_by(lender_name) %>% summarise("day-6"=n())

x <- format(sub-4,"%b-%d")
names(Day_6) <- c("lender_name",x)
names(Day_6_Not) <- c("lender_name",x)





Day_7_not_received <-dfCISdump %>% filter(subscription_date >= sub-6 & subscription_date < sub-5) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-7")
Day_7_not_sent <-dfCISdump %>% filter(subscription_date >= sub-6 & subscription_date < sub-5) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-7")


Day_7 <- Day_7_not_received %>% group_by(lender_name) %>% summarise("day-7"=n())

Day_7_Not <- Day_7_not_sent %>% group_by(lender_name) %>% summarise("day-7"=n())

x <- format(sub-5,"%b-%d")
names(Day_7) <- c("lender_name",x)
names(Day_7_Not) <- c("lender_name",x)




Day_15_not_received <-dfCISdump %>% filter(subscription_date >= sub-14 & subscription_date < sub-6) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Day-8 to 15")
Day_15_not_sent <-dfCISdump %>% filter(subscription_date >= sub-14 & subscription_date < sub-6) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Day-8 to 15")


Day_15 <- Day_15_not_received %>% group_by(lender_name) %>% summarise("day-(8-15)"=n())

Day_15_Not <- Day_15_not_sent %>% group_by(lender_name) %>% summarise("day-(8-15)"=n())

x <- paste(format(sub-6,"%b-%d")," To ",format(sub-14,"%b-%d"),sep = '')
names(Day_15) <- c("lender_name",x)
names(Day_15_Not) <- c("lender_name",x)


Beyond_15_not_received <-dfCISdump %>% filter(subscription_date < sub-14) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus != "00 To Start Action")%>% mutate(Flag="Beyond 15 Days")
Beyond_15_not_sent <-dfCISdump %>% filter(subscription_date < sub-14) %>% filter(prodstatus %in% c("CIS-WP-LP-310","CIS-WP-LD-220")) %>% 
  filter(Facilitystatus == "00 To Start Action")%>% mutate(Flag="Beyond 15 Days")


Beyond_15 <- Beyond_15_not_received %>% group_by(lender_name) %>% summarise("Beyond 15 Days"=n())

Beyond_15_Not <- Beyond_15_not_sent %>% group_by(lender_name) %>% summarise("Beyond 15 Days"=n())

x <- paste(format(sub-15,"%b-%d")," To end",sep = '')
names(Beyond_15) <- c("lender_name",x)
names(Beyond_15_Not) <- c("lender_name",x)




Not_Received <- Reduce(function(...) merge(..., all = TRUE,by = "lender_name"),
                       list(Day_1, Day_2, Day_3,Day_4,Day_5,Day_6,Day_7,Day_15,Beyond_15))

Not_Received[is.na(Not_Received)] <- 0

Not_Sent <- Reduce(function(...) merge(..., all = TRUE, by = "lender_name"),
                   list(Day_1_Not, Day_2_Not, Day_3_Not,Day_4_Not,Day_5_Not,Day_6_Not,Day_7_Not,Day_15_Not,Beyond_15_Not))

Not_Sent[is.na(Not_Sent)] <- 0

temp <- Not_Received %>% select(-lender_name)
temp$Total <- apply(temp,1,sum)
temp <- bind_rows(temp, apply(temp,2,sum))


`Lender Name` <- Not_Received$lender_name
`Lender Name` <- `Lender Name` %>% append('Total')


Not_Received_final <- cbind(`Lender Name`, temp)


temp1 <- Not_Sent %>% select(-lender_name)
temp1$Total <- apply(temp1,1,sum)
temp1 <- bind_rows(temp1, apply(temp1,2,sum))


`Lender Name` <- Not_Sent$lender_name
`Lender Name` <- `Lender Name` %>% append('Total')


Not_Sent_final <- cbind(`Lender Name`, temp1)


Not_received_File <- rbind(Day_1_not_received,Day_2_not_received,Day_3_not_received,Day_4_not_received,Day_5_not_received,
                           Day_6_not_received,Day_7_not_received,Day_15_not_received,Beyond_15_not_received)



Not_sent_File <- rbind(Day_1_not_sent,Day_2_not_sent,Day_3_not_sent,Day_4_not_sent,Day_5_not_sent,Day_6_not_sent,
                       Day_7_not_sent,Day_15_not_sent,Beyond_15_not_sent)



path <- paste("./Output/",today,"/Revert Not received from lender.xlsx",sep = "")
FileCreate(dataset=Not_Received_final,sheet_name="NOT RECEIVED MIS",path,
           sheet_name2 = "Leads",dataset2 = Not_received_File)

path1 <- paste("./Output/",today,"/Leads not sent to lender.xlsx.xlsx",sep = "")
FileCreate(dataset=Not_Sent_final,sheet_name="NOT SENT MIS",path1,
           sheet_name2 = "Leads",dataset2 = Not_sent_File)


x <-xtable(Not_Received_final,digits = 0)
y <-xtable(Not_Sent_final,digits = 0)


message <- str_interp("Revert Not received from lender:")
message1 <- str_interp("Leads not sent to lender:")


lapply(Count_Mail$`File Name`,function(z){
  #browser()
  # sender <- "referrals@creditmantri.com"
  sender <- 'Ops-cis@creditmantri.com'
  recepients <- str_split(Count_Mail$To[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  myMessage <- paste(Count_Mail$Subject[Count_Mail$`File Name` == z]," ",Sys.Date(),sep = '')
  cclist = str_split(Count_Mail$cc[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  msg <- paste('
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<style>
table {font-family:  Verdana, Geneva, sans-serif; font-size: 11px;
width = 100%; border: 1px solid black; border-collapse: collapse;
text-align: center; padding: 5px;}
th {height = 12px;background-color: #4CAF50;color: white;}
td {background-color: #FFF;}
#red{
  background-color:red;
}
</style>
</head>
<body>
  <h3>${message}</h3> 
  <p>${print(x,type =\'html\', 
      hline.after=-1:2,
      sanitize.text.function = function(x) x)} </p>
      
    <h3>${message1}</h3> 
  <p>${print(y,type =\'html\', 
      hline.after=-1:2,
      sanitize.text.function = function(x) x)} </p>    
  
  </body>
</html>')
  
  filename <-c(path,path1)
  if(z == "DAILY NEW CASE BATCH DATA REPORT"){
    mail(sender,recepients,myMessage,msg,cclist,filename)
  }
})



#Nldb------------------------------------------------------------------------------------

NLDB <- paste('.\\Output\\',today,'\\NLDB',sep = '')
if(!dir.exists(NLDB)){
  dir.create(NLDB)
}

wh <- fread('./Input/CIS Dump_WH_Dump.csv')

wh_dump <- wh %>% select(phone_home,
                         first_name,
                         last_name,
                         lead_id,
                         sub_month_orig,
                         asset_classification,
                         subscription_date,
                         lender,
                         is_ldb,
                         product,
                         account_status,
                         facilitystatus,
                         product_status,
                         Account_No) %>% filter(product_status %in% c("CIS-WP-LP-310","CIS-WP-LD-220","CIS-WP-LP-325"))

wh_dump$subscription_date <-dmy_hm(wh_dump$subscription_date) 

Mo_to_M4 <- wh_dump %>% filter(sub_month_orig %in% c("M0","M1","M2","M3","M4"))

M4_to_Missing <- wh_dump %>% filter(!sub_month_orig %in% c("M0","M1","M2","M3","M4"))



Filename <- paste('Output/',today,'/NLDB/M0 to M4 DUMP(New cases).xlsx',sep = '')
FileCreate(dataset=Mo_to_M4,sheet_name="Priority Customers",Filename)

Filename <- paste('Output/',today,'/NLDB/M4+ to Missing DUMP(LTD Cases).xlsx',sep = '')
FileCreate(dataset=M4_to_Missing,sheet_name="LTD Cases",Filename)



lapply(Count_Mail$`File Name`,function(z){
  #browser()
  # sender <- "referrals@creditmantri.com"
  sender <- 'Ops-cis@creditmantri.com'
  recepients <- str_split(Count_Mail$To[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  myMessage <- paste(Count_Mail$Subject[Count_Mail$`File Name` == z]," ",Sys.Date(),sep = '')
  cclist = str_split(Count_Mail$cc[Count_Mail$`File Name` == z], ",\\s*")[[1]]
  msg <- paste("Hi All,<br><br>PFA.<br><br>Regards<br>Credit Mantri")
  filename <- c(paste(".\\Output\\",today,"\\NLDB\\M4+ to Missing DUMP(LTD Cases).xlsx",sep = ""),paste(".\\Output\\",today,"\\NLDB\\M0 to M4 DUMP(New cases).xlsx",sep = ""))
  
  if(z == "NLDB BATCH"){
    mail(sender,recepients,myMessage,msg,cclist,filename)
  }
})










