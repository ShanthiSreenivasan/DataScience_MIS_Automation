setwd("F:/IPAC/Chennai Telephony/daily_and_hourly_trackers/daily trackers/")
library(readxl)
library(dplyr)
library(data.table)
library(lubridate)
library(openxlsx)

track_date = "mar_01"
########################## FEB 08 data ########################################
attendance_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/Attendance 08.02.2021.xlsx", sheet = 2)
attendance_feb_08 <- attendance_feb_08[,1:6]
names(attendance_feb_08) <- c("s_no", "agent_id", "agent_name", "tl_name", "wfh_status", "calling_type")
attendance_feb_08$s_no <- NULL
attendance_feb_08$agent_id <- as.character(attendance_feb_08$agent_id)
attendance_feb_08$calling_type <- tolower(attendance_feb_08$calling_type)
attendance_feb_08$tl_name <- toupper(attendance_feb_08$tl_name)
attendance_feb_08$wfh_status <- toupper(attendance_feb_08$wfh_status)
attendance_feb_08 <- attendance_feb_08[!(is.na(attendance_feb_08$agent_id)),]

training_agents_feb_08 <- attendance_feb_08[attendance_feb_08$tl_name %like% "TRAINING",]
gen_survey_agents_feb_08 <- attendance_feb_08[(!attendance_feb_08$tl_name %like% "TRAINING")&(attendance_feb_08$calling_type %in% c("tn survey", "general survey", "genral survey recalling", "survey re-calling")),]
ab_survey_agents_feb_08 <- attendance_feb_08[ attendance_feb_08$calling_type %like% "ab testing" | 
                                  attendance_feb_08$calling_type %like% "a/b testing" |
                                  attendance_feb_08$calling_type %like% "candidate survey" | 
                                  attendance_feb_08$calling_type %like% "candidate general",]

attendance_feb_08 <- attendance_feb_08[(!attendance_feb_08$calling_type %in% c("inbound", "mks outreach 2")) & 
                           (! attendance_feb_08$tl_name %like% "TRAINING") & !is.na(attendance_feb_08$calling_type),]

##########################  Gen_Survey
dispo_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/Disposition report 08.02.2021.xlsx", sheet = 1)
dispo_feb_08$CampName <- tolower(dispo_feb_08$CampName)
gen_dispo_feb_08 <- dispo_feb_08[dispo_feb_08$CampName %like% "re-calling",]
names(gen_dispo_feb_08)[3] <- "agent_id" 

######## dialed 
gen_dialed_eod_feb_08 <- gen_dispo_feb_08 %>% group_by(agent_id) %>% summarise(total_dialed = length(custno))
gen_dialed_eod_feb_08 <- merge(attendance_feb_08[,c("agent_id","wfh_status")], gen_dialed_eod_feb_08, by = "agent_id", all.x = T)

gen_team_dialed_eod_feb_08 <- merge(attendance_feb_08[,c("agent_id","tl_name")], gen_dialed_eod_feb_08, by = "agent_id", all.y = T)
gen_team_dialed_eod_feb_08 <- gen_team_dialed_eod_feb_08 %>% group_by(tl_name) %>% summarise(total_dialed = sum(total_dialed))
names(gen_dialed_eod_feb_08)[3] <- "feb_08"
names(gen_team_dialed_eod_feb_08)[2] <- "feb_08"

######## total, valid responses, conversion
gen_resp_feb_08 <- read.csv("C:/Users/Vivek/Downloads/feb_08_reports/agent_hourly.csv")
names(gen_resp_feb_08)[2] <- "agent_id"
gen_resp_feb_08$conversion <- round(gen_resp_feb_08$Valid.response/gen_resp_feb_08$Total.response*100,2) 
######## total
gen_total_resp_feb_08 <- gen_resp_feb_08[,c("agent_id", "Total.response")]
gen_total_resp_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], gen_total_resp_feb_08, by = "agent_id", all.x = T)

gen_team_total_resp_feb_08 <- merge(gen_total_resp_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_total_resp_feb_08 <- gen_team_total_resp_feb_08 %>% group_by(tl_name) %>% 
                              summarise(Total.response = sum(Total.response, na.rm = T))
names(gen_total_resp_feb_08)[3] <- "feb_08"
names(gen_team_total_resp_feb_08)[2] <- "feb_08"

######## valid
gen_valid_resp_feb_08 <- gen_resp_feb_08[,c("agent_id", "Valid.response")]
gen_valid_resp_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], gen_valid_resp_feb_08, by = "agent_id", all.x = T)

gen_team_valid_resp_feb_08 <- merge(gen_valid_resp_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_valid_resp_feb_08 <- gen_team_valid_resp_feb_08 %>% group_by(tl_name) %>% 
  summarise(Valid.response = sum(Valid.response,na.rm = T))
names(gen_valid_resp_feb_08)[3] <- "feb_08"
names(gen_team_valid_resp_feb_08)[2] <- "feb_08"

######## conversion
gen_total_to_valid_feb_08 <- gen_resp_feb_08[,c("agent_id", "conversion")]
gen_total_to_valid_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], gen_total_to_valid_feb_08, by = "agent_id", all.x = T)

gen_team_total_to_valid_feb_08 <- merge(gen_total_to_valid_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_total_to_valid_feb_08 <- gen_team_total_to_valid_feb_08 %>% group_by(tl_name) %>% 
  summarise(convrsion = mean(conversion,na.rm = T))
names(gen_total_to_valid_feb_08)[3] <- "feb_08"
names(gen_team_total_to_valid_feb_08)[2] <- "feb_08"

######### Accuracy
gen_qc_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/TN General Survey QC Report.xlsx", sheet = 2)
names(gen_qc_feb_08) <- tolower(names(gen_qc_feb_08))

######### Question Accuracy
gen_ques_accuracy_feb_08 <- gen_qc_feb_08[,c("agent_id", "question_level_accuracy")]
gen_ques_accuracy_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], gen_ques_accuracy_feb_08, by = "agent_id", all.x = T)

gen_team_ques_acc_feb_08 <- merge(attendance_feb_08[,c("agent_id", "tl_name")], gen_ques_accuracy_feb_08,  by = "agent_id", all.y = T)
gen_team_ques_acc_feb_08 <- gen_team_ques_acc_feb_08 %>% group_by(tl_name) %>%
                            summarise(question_level_accuracy = mean(question_level_accuracy, na.rm = T))

names(gen_ques_accuracy_feb_08)[3] <- "feb_08"
names(gen_team_ques_acc_feb_08)[2] <- "feb_08"

######### call Accuracy
gen_call_accuracy_feb_08 <- gen_qc_feb_08[,c("agent_id", "call_level_accuracy")]
gen_call_accuracy_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], gen_call_accuracy_feb_08, by = "agent_id", all.x = T)

gen_team_call_acc_feb_08 <- merge(gen_call_accuracy_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_call_acc_feb_08 <- gen_team_call_acc_feb_08 %>% group_by(tl_name) %>%
  summarise(call_level_accuracy = mean(call_level_accuracy, na.rm = T))

names(gen_call_accuracy_feb_08)[3] <- "feb_08"
names(gen_team_call_acc_feb_08)[2] <- "feb_08"

############### General agent level workbook 

gen_agent_wb <- createWorkbook("gen_day_to_day_agent_tracker")
addWorksheet(gen_agent_wb, sheetName = "question_level_accuracy")
writeData(gen_agent_wb, "question_level_accuracy", gen_ques_accuracy_feb_08, rowNames = F)

addWorksheet(gen_agent_wb, sheetName = "call_level_accuracy")
writeData(gen_agent_wb, "call_level_accuracy", gen_call_accuracy_feb_08, rowNames = F)

addWorksheet(gen_agent_wb, sheetName = "dialed_numbers")
writeData(gen_agent_wb, "dialed_numbers", gen_dialed_eod_feb_08, rowNames = F)

addWorksheet(gen_agent_wb, sheetName = "total_response")
writeData(gen_agent_wb, "total_response", gen_total_resp_feb_08, rowNames = F)

addWorksheet(gen_agent_wb, sheetName = "valid_response")
writeData(gen_agent_wb, "valid_response", gen_valid_resp_feb_08, rowNames = F)

addWorksheet(gen_agent_wb, sheetName = "total_to_valid_conversion")
writeData(gen_agent_wb, "total_to_valid_conversion", gen_total_to_valid_feb_08, rowNames = F)

saveWorkbook(gen_agent_wb, file = "gen_agent_level_tracker.xlsx", overwrite = T)

############### General team level workbook 

gen_team_wb <- createWorkbook("gen_day_to_day_team_tracker")
addWorksheet(gen_team_wb, sheetName = "question_level_accuracy")
writeData(gen_team_wb, "question_level_accuracy", gen_team_ques_acc_feb_08, rowNames = F)

addWorksheet(gen_team_wb, sheetName = "call_level_accuracy")
writeData(gen_team_wb, "call_level_accuracy", gen_team_call_acc_feb_08, rowNames = F)

addWorksheet(gen_team_wb, sheetName = "dialed_numbers")
writeData(gen_team_wb, "dialed_numbers", gen_team_dialed_eod_feb_08, rowNames = F)

addWorksheet(gen_team_wb, sheetName = "total_response")
writeData(gen_team_wb, "total_response", gen_team_total_resp_feb_08, rowNames = F)

addWorksheet(gen_team_wb, sheetName = "valid_response")
writeData(gen_team_wb, "valid_response", gen_team_valid_resp_feb_08, rowNames = F)

addWorksheet(gen_team_wb, sheetName = "total_to_valid_conversion")
writeData(gen_team_wb, "total_to_valid_conversion", gen_team_total_to_valid_feb_08, rowNames = F)

saveWorkbook(gen_team_wb, file = "gen_team_level_tracker.xlsx", overwrite = T)

######## AB Survey Feb 5
ab_dispo_feb_08 <- dispo_feb_08[dispo_feb_08$CampName %like% "ab testing",]
names(ab_dispo_feb_08)[3] <- "agent_id" 

######## dialed 
ab_dialed_eod_feb_08 <- ab_dispo_feb_08 %>% group_by(agent_id) %>% summarise(total_dialed = length(custno))
ab_dialed_eod_feb_08 <- merge(attendance_feb_08[,c("agent_id","wfh_status")], ab_dialed_eod_feb_08, by = "agent_id", all.x = T)

ab_team_dialed_eod_feb_08 <- merge(attendance_feb_08[,c("agent_id","tl_name")], ab_dialed_eod_feb_08, by = "agent_id", all.y = T)
ab_team_dialed_eod_feb_08 <- ab_team_dialed_eod_feb_08 %>% group_by(tl_name) %>% summarise(total_dialed = sum(total_dialed))
names(ab_dialed_eod_feb_08)[3] <- "feb_08"
names(ab_team_dialed_eod_feb_08)[2] <- "feb_08"

######## total, valid responses, conversion
ab_resp_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/TN_Summary_AB2021-02-08.xlsx", sheet = 3)
ab_resp_feb_08 <- ab_resp_feb_08[,c("agent_id", "total_responses")]

######## total
ab_total_resp_feb_08 <- ab_resp_feb_08[,c("agent_id", "total_responses")]
ab_total_resp_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], ab_total_resp_feb_08, by = "agent_id", all.x = T)

ab_team_total_resp_feb_08 <- merge(ab_total_resp_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_total_resp_feb_08 <- ab_team_total_resp_feb_08 %>% group_by(tl_name) %>% 
  summarise(total_responses = sum(total_responses, na.rm = T))
names(ab_total_resp_feb_08)[3] <- "feb_08"
names(ab_team_total_resp_feb_08)[2] <- "feb_08"

######## valid
ab_resp_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/TN_Summary_AB2021-02-08.xlsx", sheet = 4)
ab_resp_feb_08 <- ab_resp_feb_08[,c("agent_id", "total_valid")]

ab_valid_resp_feb_08 <- ab_resp_feb_08[,c("agent_id", "total_valid")]
ab_valid_resp_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], ab_valid_resp_feb_08, by = "agent_id", all.x = T)

ab_team_valid_resp_feb_08 <- merge(ab_valid_resp_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_valid_resp_feb_08 <- ab_team_valid_resp_feb_08 %>% group_by(tl_name) %>% 
  summarise(total_valid = sum(total_valid,na.rm = T))
names(ab_valid_resp_feb_08)[3] <- "feb_08"
names(ab_team_valid_resp_feb_08)[2] <- "feb_08"

######## conversion
ab_total_to_valid_feb_08 <- merge(ab_total_resp_feb_08, ab_valid_resp_feb_08, by = c("agent_id", "wfh_status"))
ab_total_to_valid_feb_08$conversion <- round(ab_total_to_valid_feb_08$feb_08.y/ab_total_to_valid_feb_08$feb_08.x*100,2)
ab_total_to_valid_feb_08 <- ab_total_to_valid_feb_08[,c("agent_id", "wfh_status", "conversion")]

ab_team_total_to_valid_feb_08 <- merge(ab_total_to_valid_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_total_to_valid_feb_08 <- ab_team_total_to_valid_feb_08 %>% group_by(tl_name) %>% 
  summarise(convrsion = mean(conversion,na.rm = T))
names(ab_total_to_valid_feb_08)[3] <- "feb_08"
names(ab_team_total_to_valid_feb_08)[2] <- "feb_08"

######### Accuracy
ab_qc_feb_08 <- read_excel("C:/Users/Vivek/Downloads/feb_08_reports/TN AB candidate QC Report.xlsx", sheet = 2)
names(ab_qc_feb_08) <- tolower(names(ab_qc_feb_08))

######### Question Accuracy
ab_ques_accuracy_feb_08 <- ab_qc_feb_08[,c("agent_id", "question_level_accuracy")]
ab_ques_accuracy_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], ab_ques_accuracy_feb_08, by = "agent_id", all.x = T)

ab_team_ques_acc_feb_08 <- merge(ab_ques_accuracy_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_ques_acc_feb_08 <- ab_team_ques_acc_feb_08 %>% group_by(tl_name) %>%
  summarise(question_level_accuracy = mean(question_level_accuracy, na.rm = T))

names(ab_ques_accuracy_feb_08)[3] <- "feb_08"
names(ab_team_ques_acc_feb_08)[2] <- "feb_08"

######### call Accuracy
ab_call_accuracy_feb_08 <- ab_qc_feb_08[,c("agent_id", "call_level_accuracy")]
ab_call_accuracy_feb_08 <- merge(attendance_feb_08[,c("agent_id", "wfh_status")], ab_call_accuracy_feb_08, by = "agent_id", all.x = T)

ab_team_call_acc_feb_08 <- merge(ab_call_accuracy_feb_08, attendance_feb_08[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_call_acc_feb_08 <- ab_team_call_acc_feb_08 %>% group_by(tl_name) %>%
  summarise(call_level_accuracy = mean(call_level_accuracy, na.rm = T))

names(ab_call_accuracy_feb_08)[3] <- "feb_08"
names(ab_team_call_acc_feb_08)[2] <- "feb_08"

############### General agent level workbook 

ab_agent_wb <- createWorkbook("ab_day_to_day_agent_tracker")
addWorksheet(ab_agent_wb, sheetName = "question_level_accuracy")
writeData(ab_agent_wb, "question_level_accuracy", ab_ques_accuracy_feb_08, rowNames = F)

addWorksheet(ab_agent_wb, sheetName = "call_level_accuracy")
writeData(ab_agent_wb, "call_level_accuracy", ab_call_accuracy_feb_08, rowNames = F)

addWorksheet(ab_agent_wb, sheetName = "dialed_numbers")
writeData(ab_agent_wb, "dialed_numbers", ab_dialed_eod_feb_08, rowNames = F)

addWorksheet(ab_agent_wb, sheetName = "total_response")
writeData(ab_agent_wb, "total_response", ab_total_resp_feb_08, rowNames = F)

addWorksheet(ab_agent_wb, sheetName = "valid_response")
writeData(ab_agent_wb, "valid_response", ab_valid_resp_feb_08, rowNames = F)

addWorksheet(ab_agent_wb, sheetName = "total_to_valid_conversion")
writeData(ab_agent_wb, "total_to_valid_conversion", ab_total_to_valid_feb_08, rowNames = F)

saveWorkbook(ab_agent_wb, file = "ab_agent_level_tracker.xlsx", overwrite = T)

############### ab team level workbook 

ab_team_wb <- createWorkbook("ab_day_to_day_team_tracker")
addWorksheet(ab_team_wb, sheetName = "question_level_accuracy")
writeData(ab_team_wb, "question_level_accuracy", ab_team_ques_acc_feb_08, rowNames = F)

addWorksheet(ab_team_wb, sheetName = "call_level_accuracy")
writeData(ab_team_wb, "call_level_accuracy", ab_team_call_acc_feb_08, rowNames = F)

addWorksheet(ab_team_wb, sheetName = "dialed_numbers")
writeData(ab_team_wb, "dialed_numbers", ab_team_dialed_eod_feb_08, rowNames = F)

addWorksheet(ab_team_wb, sheetName = "total_response")
writeData(ab_team_wb, "total_response", ab_team_total_resp_feb_08, rowNames = F)

addWorksheet(ab_team_wb, sheetName = "valid_response")
writeData(ab_team_wb, "valid_response", ab_team_valid_resp_feb_08, rowNames = F)

addWorksheet(ab_team_wb, sheetName = "total_to_valid_conversion")
writeData(ab_team_wb, "total_to_valid_conversion", ab_team_total_to_valid_feb_08, rowNames = F)

saveWorkbook(ab_team_wb, file = "ab_team_level_tracker.xlsx", overwrite = T)

#####################################################################################################################
  
attendance <- read_excel("Attendance 18.02.2021.xlsx", sheet = 1)
attendance <- attendance[,1:6]
names(attendance) <- c("s_no", "agent_id", "agent_name", "tl_name", "wfh_status", "calling_type")
attendance$s_no <- NULL
attendance$agent_id <- as.character(attendance$agent_id)
attendance$calling_type <- tolower(attendance$calling_type)
attendance$tl_name <- toupper(attendance$tl_name)
attendance$wfh_status <- toupper(attendance$wfh_status)
attendance <- attendance[!(is.na(attendance$agent_id)),]

training_agents <- attendance[attendance$tl_name %like% "TRAINING",]
gen_survey_agents <- attendance[(!attendance$tl_name %like% "TRAINING")&(attendance$calling_type %in% c("general survey", "genral survey recalling", "survey re-calling")),]
ab_survey_agents <- attendance[ attendance$calling_type %like% "ab testing" | 
                                  attendance$calling_type %like% "a/b testing" |
                                  attendance$calling_type %like% "candidate survey" | 
                                  attendance$calling_type %like% "candidate general",]

attendance <- attendance[(!attendance$calling_type %in% c("inbound", "mks outreach 2")) & 
                           (! attendance$tl_name %like% "TRAINING") & !is.na(attendance$calling_type),]


################# GEN SURVEY TRACKER  ###############################################
files <- (list.files()[grepl("csv|xlsx",list.files())])
attendance_file <- list.files()[list.files() %like% "Attend"]
disposition_file <- list.files()[list.files() %like% "Dispo"|list.files() %like% "DISPO"]

dispo <- read_excel(disposition_file, sheet = 1)
dispo$CampName <- tolower(dispo$CampName)
gen_dispo <- dispo[dispo$CampName %like% "survey re-calling" | dispo$CampName %like% "tn-survey", ]
gen_dispo$agentid <- as.character(gen_dispo$agentid) 
names(gen_dispo)[3] <- "agent_id"

############## attendance #################
attendance <- read_excel(attendance_file, sheet = 1)
attendance <- attendance[,1:6]
names(attendance) <- c("s_no", "agent_id", "tl_name", "agent_name", "wfh_status", "calling_type")
attendance$s_no <- NULL
attendance$agent_id <- as.character(attendance$agent_id)
attendance$calling_type <- tolower(attendance$calling_type)
attendance$tl_name <- toupper(attendance$tl_name)
attendance$wfh_status <- toupper(attendance$wfh_status)
attendance <- attendance[!(is.na(attendance$agent_id)),]

training_agents <- attendance[attendance$tl_name %like% "TRAINING",]
gen_survey_agents <- attendance[(!attendance$tl_name %like% "TRAINING")&(attendance$calling_type %in% c("tn survey", "general survey", "genral survey recalling", "survey re-calling", "tn genral")),]
ab_survey_agents <- attendance[ attendance$calling_type %like% "ab testing" | 
                                  attendance$calling_type %like% "a/b testing" |
                                  attendance$calling_type %like% "candidate survey" | 
                                  attendance$calling_type %like% "candidate general",]

attendance <- attendance[(!attendance$calling_type %in% c("inbound", "mks outreach 2")) & 
                           (! attendance$tl_name %like% "TRAINING") & !is.na(attendance$calling_type),]

########## dialed numbers

gen_dialed_eod <- gen_dispo %>% group_by(agent_id) %>% summarise(total_dialed = length(custno))
gen_dialed_eod <- merge(gen_survey_agents[,c("agent_id", "wfh_status")], gen_dialed_eod, by = "agent_id", all.x = T)
gen_team_dialed_eod <- merge(gen_dialed_eod, gen_survey_agents[,c("agent_id", "tl_name")], all.x = T)
gen_team_dialed_eod <- gen_team_dialed_eod %>% group_by(tl_name) %>% summarise(team_dialed = sum(total_dialed))
names(gen_dialed_eod)[3] <- track_date
names(gen_team_dialed_eod)[2] <- track_date

########## Accuracy
gen_qc <- read_excel("TN General Survey QC Report.xlsx", sheet = 2)
names(gen_qc) <- tolower(names(gen_qc))
gen_qc <- merge(gen_qc, gen_survey_agents[,c("agent_id", "wfh_status")], all.y=T)
gen_qc_wfh <- gen_qc %>% group_by(wfh_status) %>% summarise(avg_accuracy = mean(question_level_accuracy, na.rm = T))
gen_qc_wfh <- gen_qc_wfh[order(gen_qc_wfh$avg_accuracy, decreasing = T),]

gen_qc_team <- merge(gen_qc, gen_survey_agents[,c("agent_id", "tl_name")], all.y = T)
gen_qc_team <- gen_qc_team %>% group_by(tl_name) %>% summarise(avg_accuracy = round(mean(question_level_accuracy, na.rm = T),2))
gen_qc_team <- gen_qc_team[order(gen_qc_team$avg_accuracy, decreasing = T),]

########### Question level accuracy
gen_qc_ques_acc <- gen_qc[,c("agent_id", "wfh_status", "question_level_accuracy")]
gen_team_ques_acc <- merge(gen_qc_ques_acc, gen_survey_agents[,c("agent_id", "tl_name")], all.y = T)
gen_team_ques_acc <- gen_team_ques_acc %>% group_by(tl_name) %>% summarise(question_level_accuracy = round(mean(question_level_accuracy, na.rm = T),2))
gen_team_ques_acc <- gen_team_ques_acc[!is.na(gen_team_ques_acc$tl_name),]
names(gen_qc_ques_acc)[3] <- track_date 
names(gen_team_ques_acc)[2] <- track_date 

########### Call level accuracy
gen_qc_call_acc <- gen_qc[,c("agent_id", "wfh_status", "call_level_accuracy")]
gen_team_call_acc <- merge(gen_qc_call_acc, gen_survey_agents[,c("agent_id", "tl_name")], all.y = T)
gen_team_call_acc <- gen_team_call_acc %>% group_by(tl_name) %>% summarise(call_level_accuracy = round(mean(call_level_accuracy, na.rm = T),2))
gen_team_call_acc <- gen_team_call_acc[!is.na(gen_team_call_acc$tl_name),]
names(gen_qc_call_acc)[3] <- track_date 
names(gen_team_call_acc)[2] <- track_date 

########### Responses
gen_first_last <- read.csv("agent_hourly.csv")
gen_first_last$X <- NULL
names(gen_first_last)[names(gen_first_last) == "agent.id"] <- "agent_id"
gen_first_last$agent_id <- as.character(gen_first_last$agent_id)
gen_first_last$First.Response <- as.character(gen_first_last$First.Response)
gen_first_last$Last.Response <- as.character(gen_first_last$Last.Response)
gen_first_last$agent_id <- gsub("^\\s|\\s$","", gen_first_last$agent_id) 

gen_responses <- gen_first_last %>% group_by(agent_id) %>% summarise(Total.response = sum(Total.response),
                                                                     Valid.response = sum(Valid.response))
gen_first_last <- gen_first_last %>% group_by(agent_id) %>% 
  arrange(First.Response) %>% mutate(rank = 1:length(First.Response)) %>%
  filter(rank == 1) %>% select(agent_id, First.Response,Last.Response)

gen_responses$conversion <- round(gen_responses$Valid.response/gen_responses$Total.response*100,2)
########### Total Responses
gen_total_resp <- gen_responses[,c("agent_id","Total.response")] 
gen_total_resp <- merge(gen_survey_agents[,c("agent_id", "wfh_status")],gen_total_resp, by = "agent_id", all.x = T)
gen_team_total_resp <- merge(gen_total_resp, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_total_resp <- gen_team_total_resp %>% group_by(tl_name) %>% 
  summarise(total_resp = sum(Total.response, na.rm = T))
gen_team_total_resp <- gen_team_total_resp[!is.na(gen_team_total_resp$tl_name),]
names(gen_total_resp)[3] <- track_date
names(gen_team_total_resp)[2] <- track_date


########### Valid Responses
gen_valid_resp <- gen_responses[,c("agent_id","Valid.response")] 
gen_valid_resp <- merge(gen_survey_agents[,c("agent_id", "wfh_status")],gen_valid_resp, by = "agent_id", all.x = T)
gen_team_valid_resp <- merge(gen_valid_resp, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_valid_resp <- gen_team_valid_resp %>% group_by(tl_name) %>% 
  summarise(valid_resp = sum(Valid.response, na.rm = T))
gen_team_valid_resp <- gen_team_valid_resp[!is.na(gen_team_valid_resp$tl_name),]
names(gen_valid_resp)[3] <- track_date
names(gen_team_valid_resp)[2] <- track_date

########### Total to Valid
gen_total_to_valid <- gen_responses[,c("agent_id", "conversion")] 
gen_total_to_valid <- merge(gen_survey_agents[,c("agent_id", "wfh_status")],gen_total_to_valid, by = "agent_id", all.x = T)
gen_team_total_to_valid <- merge(gen_total_to_valid, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_team_total_to_valid <- gen_team_total_to_valid %>% group_by(tl_name) %>% 
  summarise(conversion = mean(conversion, na.rm = T))
gen_team_total_to_valid <- gen_team_total_to_valid[!is.na(gen_team_total_to_valid$tl_name),]
names(gen_total_to_valid)[3] <- track_date
names(gen_team_total_to_valid)[2] <- track_date

############# Gen Survey Agent level workbook
gen_agent_wb <- createWorkbook("gen_day_to_day_agent_tracker")

daily_gen_agent_ques_acc <- read_excel("gen_agent_level_tracker.xlsx", sheet = 1)
daily_gen_agent_ques_acc <- merge(daily_gen_agent_ques_acc, gen_qc_ques_acc, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_ques_acc <- unique(daily_gen_agent_ques_acc)
addWorksheet(gen_agent_wb, sheetName = "question_level_accuracy")
writeData(gen_agent_wb, "question_level_accuracy", daily_gen_agent_ques_acc, rowNames = F)

daily_gen_agent_call_acc <- read_excel("gen_agent_level_tracker.xlsx", sheet = 2)
daily_gen_agent_call_acc <- merge(daily_gen_agent_call_acc, gen_qc_call_acc, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_call_acc <- unique(daily_gen_agent_call_acc)
addWorksheet(gen_agent_wb, sheetName = "call_level_accuracy")
writeData(gen_agent_wb, "call_level_accuracy", daily_gen_agent_call_acc, rowNames = F)

daily_gen_agent_dialed <- read_excel("gen_agent_level_tracker.xlsx", sheet = 3)
daily_gen_agent_dialed <- merge(daily_gen_agent_dialed, gen_dialed_eod, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_dialed <- unique(daily_gen_agent_dialed)
addWorksheet(gen_agent_wb, sheetName = "survey_dialed_numbers")
writeData(gen_agent_wb, "survey_dialed_numbers", daily_gen_agent_dialed, rowNames = F)

daily_gen_agent_total_resp <- read_excel("gen_agent_level_tracker.xlsx", sheet = 4)
daily_gen_agent_total_resp <- merge(daily_gen_agent_total_resp, gen_total_resp, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_total_resp <- unique(daily_gen_agent_total_resp)
addWorksheet(gen_agent_wb, sheetName = "total_response")
writeData(gen_agent_wb, "total_response", daily_gen_agent_total_resp, rowNames = F)

daily_gen_agent_valid_resp <- read_excel("gen_agent_level_tracker.xlsx", sheet = 5)
daily_gen_agent_valid_resp <- merge(daily_gen_agent_valid_resp, gen_valid_resp, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_valid_resp <- unique(daily_gen_agent_valid_resp)
addWorksheet(gen_agent_wb, sheetName = "valid_response")
writeData(gen_agent_wb, "valid_response", daily_gen_agent_valid_resp, rowNames = F)

daily_gen_agent_total_to_valid <- read_excel("gen_agent_level_tracker.xlsx", sheet = 6)
daily_gen_agent_total_to_valid <- merge(daily_gen_agent_total_to_valid, gen_total_to_valid, by = c("agent_id", "wfh_status"), all = T)
daily_gen_agent_total_to_valid <- unique(daily_gen_agent_total_to_valid)
addWorksheet(gen_agent_wb, sheetName = "total_to_valid_conversion")
writeData(gen_agent_wb, "total_to_valid_conversion", daily_gen_agent_total_to_valid, rowNames = F)

saveWorkbook(gen_agent_wb, file = "gen_agent_level_tracker.xlsx", overwrite = T)

############# Gen Survey Team level workbook

gen_team_wb <- createWorkbook("gen_day_to_day_team_tracker")

daily_gen_team_ques_acc <- read_excel("gen_team_level_tracker.xlsx", sheet = 1)
daily_gen_team_ques_acc <- merge(daily_gen_team_ques_acc, gen_team_ques_acc, by = "tl_name", all = T)
daily_gen_team_ques_acc <- unique(daily_gen_team_ques_acc)
addWorksheet(gen_team_wb, sheetName = "question_level_accuracy")
writeData(gen_team_wb, "question_level_accuracy", daily_gen_team_ques_acc, rowNames = F)

daily_gen_team_call_acc <- read_excel("gen_team_level_tracker.xlsx", sheet = 2)
daily_gen_team_call_acc <- merge(daily_gen_team_call_acc, gen_team_call_acc, by = "tl_name", all = T)
daily_gen_team_call_acc <- unique(daily_gen_team_call_acc)
addWorksheet(gen_team_wb, sheetName = "call_level_accuracy")
writeData(gen_team_wb, "call_level_accuracy", daily_gen_team_call_acc, rowNames = F)

daily_gen_team_dialed <- read_excel("gen_team_level_tracker.xlsx", sheet = 3)
daily_gen_team_dialed <- merge(daily_gen_team_dialed, gen_team_dialed_eod, by = "tl_name", all = T)
daily_gen_team_dialed <- unique(daily_gen_team_dialed)
addWorksheet(gen_team_wb, sheetName = "dialed_numbers")
writeData(gen_team_wb, "dialed_numbers", daily_gen_team_dialed, rowNames = F)

daily_gen_team_total_resp <- read_excel("gen_team_level_tracker.xlsx", sheet = 4)
daily_gen_team_total_resp <- merge(daily_gen_team_total_resp, gen_team_total_resp, by = "tl_name", all = T)
daily_gen_team_total_resp <- unique(daily_gen_team_total_resp)
addWorksheet(gen_team_wb, sheetName = "total_response")
writeData(gen_team_wb, "total_response", daily_gen_team_total_resp, rowNames = F)

daily_gen_team_valid_resp <- read_excel("gen_team_level_tracker.xlsx", sheet = 5)
daily_gen_team_valid_resp <- merge(daily_gen_team_valid_resp, gen_team_valid_resp, by = "tl_name", all = T)
daily_gen_team_valid_resp <- unique(daily_gen_team_valid_resp)
addWorksheet(gen_team_wb, sheetName = "valid_response")
writeData(gen_team_wb, "valid_response", daily_gen_team_valid_resp, rowNames = F)

daily_gen_team_total_to_valid <- read_excel("gen_team_level_tracker.xlsx", sheet = 6)
daily_gen_team_total_to_valid <- merge(daily_gen_team_total_to_valid, gen_team_total_to_valid, by = "tl_name", all = T)
daily_gen_team_total_to_valid <- unique(daily_gen_team_total_to_valid)
addWorksheet(gen_team_wb, sheetName = "total_to_valid_conversion")
writeData(gen_team_wb, "total_to_valid_conversion", daily_gen_team_total_to_valid, rowNames = F)

saveWorkbook(gen_team_wb, file = "gen_team_level_tracker.xlsx", overwrite = T)

##################   AB SURVEY TRACKER #########################################

#ab_dispo <- read_excel("Disposition report 06.02.2021.xlsx", sheet = 1)
ab_dispo <- dispo[dispo$CampName %like% "candidate survey" | 
                    dispo$CampName %like% "ab testing",]
ab_dispo$agentid <- as.character(ab_dispo$agentid)
names(ab_dispo)[3] <- "agent_id"


############ dialed Numbers
ab_dialed_eod <- ab_dispo %>% group_by(agent_id) %>% summarise(total_dialed = length(custno))
ab_dialed_eod <- merge(ab_survey_agents[,c("agent_id", "wfh_status")],ab_dialed_eod, by = "agent_id", all.x = T)
ab_team_dialed_eod <- merge(ab_dialed_eod, ab_survey_agents[,c("agent_id", "tl_name")], all.x = T)
ab_team_dialed_eod <- ab_team_dialed_eod %>% group_by(tl_name) %>% summarise(total_dialed = sum(total_dialed))
names(ab_dialed_eod)[3] <- track_date
names(ab_team_dialed_eod)[2] <- track_date

############ Accuracy
# a <- read_excel("TN Candidate Survey QC Report.xlsx", sheet = 2)
# b <- read_excel("TN AB candidate QC Report.xlsx", sheet = 2)
# ab_qc <- rbind(a,b)
ab_qc_file <- files[(files %like% "AB" | files %like% "Cand") & files %like% "QC"]
ab_qc <- read_excel(ab_qc_file, sheet = 2)
names(ab_qc) <- tolower(names(ab_qc))
ab_qc <- merge(ab_qc, attendance[,c("agent_id", "wfh_status")], all.x=T)
ab_qc_wfh <- ab_qc %>% group_by(wfh_status) %>% summarise(avg_accuracy = mean(question_level_accuracy, na.rm = T))
ab_qc_wfh <- ab_qc_wfh[order(ab_qc_wfh$avg_accuracy, decreasing = T),]

ab_qc_team <- merge(ab_qc, attendance[,c("agent_id", "tl_name")], all.x = T)
ab_qc_team <- ab_qc_team %>% group_by(tl_name) %>% summarise(avg_accuracy = round(mean(question_level_accuracy, na.rm = T),2))
ab_qc_team <- ab_qc_team[order(ab_qc_team$avg_accuracy, decreasing = T),]

############ Question level accuracy
ab_ques_acc <- ab_qc[,c("agent_id", "wfh_status", "question_level_accuracy")]
ab_team_ques_acc <- merge(ab_ques_acc, ab_survey_agents[,c("agent_id", "tl_name")], all.y = T)
ab_team_ques_acc <- ab_team_ques_acc %>% group_by(tl_name) %>% summarise(question_level_accuracy = round(mean(question_level_accuracy, na.rm = T),2))
ab_team_ques_acc <- ab_team_ques_acc[!is.na(ab_team_ques_acc$tl_name),]
names(ab_team_ques_acc)[2] <- track_date 
names(ab_ques_acc)[3] <- track_date 

############ Call level accuracy
ab_call_acc <- ab_qc[,c("agent_id", "wfh_status", "call_level_accuracy")]
ab_team_call_acc <- merge(ab_call_acc, ab_survey_agents[,c("agent_id", "tl_name")], all.y = T)
ab_team_call_acc <- ab_team_call_acc %>% group_by(tl_name) %>% summarise(call_level_accuracy = round(mean(call_level_accuracy, na.rm = T),2))
ab_team_call_acc <- ab_team_call_acc[!is.na(ab_team_call_acc$tl_name),]
names(ab_team_call_acc)[2] <- track_date 
names(ab_call_acc)[3] <- track_date 

############ Agent Responses
############ Total Response
# a <- files[files %like% "Summary_Cand"]
# a <- read_excel(a, sheet = 3)
# a <- a[,c("agent_id", "total_responses")]
# b <- files[files %like% "Summary_AB"]
# b <- read_excel(b, sheet = 3)
# b <- b[,c("agent_id", "total_responses")]
# ab_total_resp <- rbind(a,b)
# ab_total_resp <- ab_total_resp %>% group_by(agent_id) %>%
#                   summarise(total_responses = sum(total_responses, na.rm = T))
responses_file <- files[files %like% "Summary"]
ab_total_resp <- read_excel(responses_file, sheet = 3)
ab_total_resp <- ab_total_resp[,c("agent_id","total_responses")]
ab_total_resp <- merge(ab_survey_agents[,c("agent_id", "wfh_status")], ab_total_resp, by = "agent_id", all.x = T)
ab_team_total_resp <- merge(ab_total_resp, ab_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_total_resp <- ab_team_total_resp %>% group_by(tl_name) %>% 
  summarise(total_resp = sum(total_responses, na.rm = T))
ab_team_total_resp <- ab_team_total_resp[!is.na(ab_team_total_resp$tl_name),]
names(ab_total_resp)[3] <- track_date
names(ab_team_total_resp)[2] <- track_date

############ valid Response
# a <- files[files %like% "Summary_Cand"]
# a <- read_excel(a, sheet = 4)
# a <- a[,c("agent_id", "total_valid")]
# b <- files[files %like% "Summary_AB"]
# b <- read_excel(b, sheet = 4)
# b <- b[,c("agent_id", "total_valid")]
# ab_valid_resp <- rbind(a,b)
# ab_valid_resp <- ab_valid_resp %>% group_by(agent_id) %>%
#                   summarise(total_valid = sum(total_valid, na.rm = T))

ab_valid_resp <- read_excel(responses_file, sheet = 4)
ab_valid_resp <- ab_valid_resp[,c("agent_id","total_valid")]
ab_valid_resp <- merge(ab_survey_agents[,c("agent_id", "wfh_status")], ab_valid_resp, by = "agent_id", all.x = T)
ab_team_valid_resp <- merge(ab_valid_resp, ab_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_team_valid_resp <- ab_team_valid_resp %>% group_by(tl_name) %>% 
  summarise(valid_resp = sum(total_valid, na.rm = T))
ab_team_valid_resp <- ab_team_valid_resp[!is.na(ab_team_valid_resp$tl_name),]
names(ab_team_valid_resp)[2] <- track_date
names(ab_valid_resp)[3] <- track_date

############ Total to valid
ab_total_to_valid <- merge(ab_total_resp, ab_valid_resp, by = c("agent_id", "wfh_status"))
ab_total_to_valid$conversion <- round(ab_total_to_valid[,4]/ab_total_to_valid[,3]*100,2)
ab_total_to_valid <- ab_total_to_valid[,c(1,2,5)]
ab_team_total_to_valid <- merge(ab_total_to_valid, ab_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.y = T)
ab_team_total_to_valid <- ab_team_total_to_valid %>% group_by(tl_name) %>% 
  summarise(conversion = mean(conversion, na.rm = T))
ab_team_total_to_valid <- ab_team_total_to_valid[!is.na(ab_team_total_to_valid$tl_name),]
names(ab_total_to_valid)[3] <- track_date
names(ab_team_total_to_valid)[2] <- track_date

############# AB Survey Agent level workbook
ab_agent_wb <- createWorkbook("ab_day_to_day_agent_tracker")
daily_ab_agent_ques_acc <- read_excel("ab_agent_level_tracker.xlsx", sheet = 1)
daily_ab_agent_ques_acc <- merge(daily_ab_agent_ques_acc, ab_ques_acc, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_ques_acc <- unique(daily_ab_agent_ques_acc)
addWorksheet(ab_agent_wb, sheetName = "question_level_accuracy")
writeData(ab_agent_wb, "question_level_accuracy", daily_ab_agent_ques_acc, rowNames = F)

daily_ab_agent_call_acc <- read_excel("ab_agent_level_tracker.xlsx", sheet = 2)
daily_ab_agent_call_acc <- merge(daily_ab_agent_call_acc, ab_call_acc, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_call_acc <- unique(daily_ab_agent_call_acc)
addWorksheet(ab_agent_wb, sheetName = "call_level_accuracy")
writeData(ab_agent_wb, "call_level_accuracy", daily_ab_agent_call_acc, rowNames = F)

daily_ab_agent_dialed <- read_excel("ab_agent_level_tracker.xlsx", sheet = 3)
daily_ab_agent_dialed <- merge(daily_ab_agent_dialed, ab_dialed_eod, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_dialed <- unique(daily_ab_agent_dialed)
addWorksheet(ab_agent_wb, sheetName = "survey_dialed_numbers")
writeData(ab_agent_wb, "survey_dialed_numbers", daily_ab_agent_dialed, rowNames = F)

daily_ab_agent_total_resp <- read_excel("ab_agent_level_tracker.xlsx", sheet = 4)
daily_ab_agent_total_resp <- merge(daily_ab_agent_total_resp, ab_total_resp, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_total_resp <- unique(daily_ab_agent_total_resp)
addWorksheet(ab_agent_wb, sheetName = "total_response")
writeData(ab_agent_wb, "total_response", daily_ab_agent_total_resp, rowNames = F)

daily_ab_agent_valid_resp <- read_excel("ab_agent_level_tracker.xlsx", sheet = 5)
daily_ab_agent_valid_resp <- merge(daily_ab_agent_valid_resp, ab_valid_resp, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_valid_resp <- unique(daily_ab_agent_valid_resp)
addWorksheet(ab_agent_wb, sheetName = "valid_response")
writeData(ab_agent_wb, "valid_response", daily_ab_agent_valid_resp, rowNames = F)

daily_ab_agent_total_to_valid <- read_excel("ab_agent_level_tracker.xlsx", sheet = 6)
daily_ab_agent_total_to_valid <- merge(daily_ab_agent_total_to_valid, ab_total_to_valid, by = c("agent_id", "wfh_status"), all = T)
daily_ab_agent_total_to_valid <- unique(daily_ab_agent_total_to_valid)
addWorksheet(ab_agent_wb, sheetName = "total_to_valid_conversion")
writeData(ab_agent_wb, "total_to_valid_conversion", daily_ab_agent_total_to_valid, rowNames = F)

saveWorkbook(ab_agent_wb, file = "ab_agent_level_tracker.xlsx", overwrite = T)

############# AB Survey Team level workbook

ab_team_wb <- createWorkbook("ab_day_to_day_team_tracker")
daily_ab_team_ques_acc <- read_excel("ab_team_level_tracker.xlsx", sheet = 1)
daily_ab_team_ques_acc <- merge(daily_ab_team_ques_acc, ab_team_ques_acc, by = "tl_name", all = T)
daily_ab_team_ques_acc <- unique(daily_ab_team_ques_acc)
addWorksheet(ab_team_wb, sheetName = "question_level_accuracy")
writeData(ab_team_wb, "question_level_accuracy", daily_ab_team_ques_acc, rowNames = F)

daily_ab_team_call_acc <- read_excel("ab_team_level_tracker.xlsx", sheet = 2)
daily_ab_team_call_acc <- merge(daily_ab_team_call_acc, ab_team_call_acc, by = "tl_name", all = T)
daily_ab_team_call_acc <- unique(daily_ab_team_call_acc)
addWorksheet(ab_team_wb, sheetName = "call_level_accuracy")
writeData(ab_team_wb, "call_level_accuracy", daily_ab_team_call_acc, rowNames = F)

daily_ab_team_dialed <- read_excel("ab_team_level_tracker.xlsx", sheet = 3)
daily_ab_team_dialed <- merge(daily_ab_team_dialed, ab_team_dialed_eod, by = "tl_name", all = T)
daily_ab_team_dialed <- unique(daily_ab_team_dialed)
addWorksheet(ab_team_wb, sheetName = "dialed_numbers")
writeData(ab_team_wb, "dialed_numbers", daily_ab_team_dialed, rowNames = F)

daily_ab_team_total_resp <- read_excel("ab_team_level_tracker.xlsx", sheet = 4)
daily_ab_team_total_resp <- merge(daily_ab_team_total_resp, ab_team_total_resp, by = "tl_name", all = T)
daily_ab_team_total_resp <- unique(daily_ab_team_total_resp)
addWorksheet(ab_team_wb, sheetName = "total_response")
writeData(ab_team_wb, "total_response", daily_ab_team_total_resp, rowNames = F)

daily_ab_team_valid_resp <- read_excel("ab_team_level_tracker.xlsx", sheet = 5)
daily_ab_team_valid_resp <- merge(daily_ab_team_valid_resp, ab_team_valid_resp, by = "tl_name", all = T)
daily_ab_team_valid_resp <- unique(daily_ab_team_valid_resp)
addWorksheet(ab_team_wb, sheetName = "valid_response")
writeData(ab_team_wb, "valid_response", daily_ab_team_valid_resp, rowNames = F)

daily_ab_team_total_to_valid <- read_excel("ab_team_level_tracker.xlsx", sheet = 6)
daily_ab_team_total_to_valid <- merge(daily_ab_team_total_to_valid, ab_team_total_to_valid, by = "tl_name", all = T)
daily_ab_team_total_to_valid <- unique(daily_ab_team_total_to_valid)
addWorksheet(ab_team_wb, sheetName = "total_to_valid_conversion")
writeData(ab_team_wb, "total_to_valid_conversion", daily_ab_team_total_to_valid, rowNames = F)

saveWorkbook(ab_team_wb, file = "ab_team_level_tracker.xlsx", overwrite = T)
