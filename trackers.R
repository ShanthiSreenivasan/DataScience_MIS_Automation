setwd("F:/IPAC/Chennai Telephony/daily_and_hourly_trackers")

library(readxl)
library(dplyr)
library(data.table)
library(lubridate)
library(openxlsx)

#################   CREATE DATE FOLDER
if(dir.exists(as.character(Sys.Date())))
{
  message("Folder exists")
} else {
  dir.create(as.character(Sys.Date()))
  message("Folder created")
}

files <- (list.files()[grepl("csv|xlsx",list.files())])
attendance_file <- list.files()[list.files() %like% "Attend" | list.files() %like% "ATTEND"]
disposition_file <- list.files()[list.files() %like% "Dispo"|list.files() %like% "DISPO"]

move_file <- function(file_name,from,to){
  file.rename( from = file.path(paste(from, file_name, sep ="/")) ,
               to = file.path(paste(to,file_name,sep ="/")))
}

#################   MOVE FILES TO DATE FOLDER
sapply(files, move_file, from = getwd(), to = paste(getwd(),Sys.Date(),sep = "/"))
setwd(paste(getwd(),Sys.Date(),sep ="/"))

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
gen_survey_agents <- attendance[(!attendance$tl_name %like% "TRAINING")&(attendance$calling_type %in% c("tn survey", "general survey", "genral survey recalling", "survey re-calling")),]
ab_survey_agents <- attendance[ attendance$calling_type %like% "ab testing" | 
                                  attendance$calling_type %like% "a/b testing" |
                                  attendance$calling_type %like% "candidate survey" | 
                                  attendance$calling_type %like% "candidate general",]

attendance <- attendance[(!attendance$calling_type %in% c("inbound", "mks outreach 2")) & 
                           (! attendance$tl_name %like% "TRAINING") & !is.na(attendance$calling_type),]

#################    MORNING ATTENDANCE CHECK
no_of_agents_per_TL <- as.data.frame(table(attendance$tl_name))
names(no_of_agents_per_TL) <- c("tl_name", "no_of_agents")
no_of_agents_per_TL$tl_name <- as.character(no_of_agents_per_TL$tl_name)

no_of_agents_per_campaign <- as.data.frame(table(attendance$calling_type))
names(no_of_agents_per_campaign) <- c("calling_type", "no_of_agents")
no_of_agents_per_campaign$calling_type <- as.character(no_of_agents_per_campaign$calling_type)

wfh_status <- as.data.frame(table(attendance$wfh_status))
names(wfh_status) <- c("wfh_status", "no_of_agents")
wfh_status$wfh_status <- as.character(wfh_status$wfh_status)

#################   CREATE HOUR FOLDER
if(dir.exists(as.character(hour(Sys.time()))))
{
  message("Folder exists")
} else {
  dir.create(as.character(hour(Sys.time())))
  message("Folder created")
}

#################   MOVE FILES TO HOUR  FOLDER
files <- files[!grepl("Attend",files)]
sapply(files, move_file, from = getwd(), to = paste(getwd(),hour(Sys.time()),sep = "/"))

#################   SET WORKING DIRECTORY FOR THE HOUR 
setwd(paste(getwd(),hour(Sys.time()),sep ="/"))


###############################     TRACK METRICS    ##################################################

disposition <- read_excel(disposition_file, sheet = 1)
disposition$CampName <- tolower(disposition$CampName)
names(disposition)[names(disposition) == "agentid"] <- "agent_id"
disposition$agent_id <- gsub("^\\s|\\s$","", disposition$agent_id) 

#################     GENERAL SURVEY TRACK  ###################################

gen_survey_qc <- read_excel("TN General Survey QC Report.xlsx", sheet = 2)
names(gen_survey_qc) <- tolower(names(gen_survey_qc))
gen_survey_qc$agent_id <- as.character(gen_survey_qc$agent_id)
gen_survey_qc$agent_id <- gsub("^\\s|\\s$","", gen_survey_qc$agent_id) 

# gen_responses <- read.csv("agent_response.csv")
# gen_responses$X <- NULL
# gen_responses$agent_id <- as.character(gen_responses$agent_id)

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
#gen_responses <- gen_first_last[,c("agent_id", "Total.response", "Valid.response")] 

gen_dispo <- disposition[disposition$CampName %in% c("survey re-calling", "tn-survey", "survey re calling"),]
gen_dispo_distrib <- as.data.frame(table(gen_dispo$Disposition))
names(gen_dispo_distrib) <- c("disposition", "frequency")
gen_dispo_distrib$disposition <- as.character(gen_dispo_distrib$disposition)
gen_dispo_distrib$total_dialed <- nrow(gen_dispo)
gen_dispo_distrib$percent <- round(gen_dispo_distrib$frequency/gen_dispo_distrib$total_dialed*100,2)  

message(paste0("general survey total number of responses = ", sum(gen_responses$total)))
message(paste0("general survey total number of complete responses = ", sum(gen_responses$complete)))
message(paste0("general survey total number of answers + auto call again = ", sum(gen_dispo_distrib$frequency[gen_dispo_distrib$disposition %in% c("Answer", "AUTO CALL AGAIN")])))

# ## agent IDs in general disposition but not in morning attendance
# gen_dialed_nums_per_agent$agentid[!(gen_dialed_nums_per_agent$agentid %in% gen_survey_agents$agent_id)]
# ## agent IDs in morning attendance but not in disposition
# gen_survey_agents$agent_id[!(gen_survey_agents$agent_id %in% gen_dialed_nums_per_agent$agentid)]

# agents_not_taking_calls <- gen_survey_agents[!gen_survey_agents$agent_id %in% gen_dispo$agent_id, ]
# agents_taking_calls_not_in_attendance <- unique(gen_dispo[!gen_dispo$agent_id %in% gen_survey_agents$agent_id, "agent_id"])
# 
# attendance_not_calling <- attendance[attendance$agent_id %in% agents_taking_calls_not_in_attendance$agent_id,]

#################   NUMBER OF CALLS 

gen_dialed_nums_per_agent <- gen_dispo %>% group_by(agent_id) %>% summarise(total_numbers_dialed = length(custno))

gen_dialed_nums_per_team <- merge(gen_dialed_nums_per_agent, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_dialed_nums_per_team <- gen_dialed_nums_per_team %>% group_by(tl_name) %>% summarise(total_numbers_dialed = sum(total_numbers_dialed))


gen_dialed_nums_per_agent$dialed_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*25, (hour(Sys.time()) - 10)*25)

gen_dialed_nums_per_team <- merge(gen_dialed_nums_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
gen_dialed_nums_per_team$dialed_target <- gen_dialed_nums_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*25, (hour(Sys.time()) - 10)*25)


names(gen_dialed_nums_per_agent)[2] <- paste0("total_dialed_",hour(Sys.time()))
names(gen_dialed_nums_per_team)[2] <- paste0("total_dialed_",hour(Sys.time()))

## agentwise
if(!exists("gen_agent_hourly_dialed_num_tracker"))
{
  gen_agent_hourly_dialed_num_tracker <- gen_dialed_nums_per_agent
} else {
  gen_agent_hourly_dialed_num_tracker$dialed_target <- NULL
  gen_agent_hourly_dialed_num_tracker <- merge(gen_agent_hourly_dialed_num_tracker, gen_dialed_nums_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("gen_team_hourly_dialed_num_tracker"))
{
  gen_team_hourly_dialed_num_tracker <- gen_dialed_nums_per_team
} else {
  gen_team_hourly_dialed_num_tracker$dialed_target <- NULL
  gen_team_hourly_dialed_num_tracker$no_of_agents <- NULL
  gen_team_hourly_dialed_num_tracker <- merge(gen_team_hourly_dialed_num_tracker, gen_dialed_nums_per_team, by = "tl_name", all = T)
}

write.csv(gen_agent_hourly_dialed_num_tracker, file = paste0("gen_agent_hourly_dialed_num_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_team_hourly_dialed_num_tracker, file = paste0("gen_team_hourly_dialed_num_tracker",hour(Sys.time()),".csv"),row.names = F)

#################   NUMBER OF FORMS

## Total Response
gen_total_response_per_agent <- gen_responses[,c("agent_id", "Total.response")]

gen_total_response_per_team <- merge(gen_total_response_per_agent, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_total_response_per_team <- gen_total_response_per_team %>% group_by(tl_name) %>% summarise(total_response = sum(Total.response))

gen_total_response_per_agent$total_response_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*5, (hour(Sys.time()) - 10)*5)

gen_total_response_per_team <- merge(gen_total_response_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
gen_total_response_per_team$total_response_target <- gen_total_response_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 10)*5, (hour(Sys.time()) - 10)*5)

names(gen_total_response_per_agent)[2] <- paste0("total_response_",hour(Sys.time()))
names(gen_total_response_per_team)[2] <- paste0("total_response_",hour(Sys.time()))

## agentwise
if(!exists("gen_agent_hourly_total_response_tracker"))
{
  gen_agent_hourly_total_response_tracker <- gen_total_response_per_agent
} else {
  gen_agent_hourly_total_response_tracker$total_response_target <- NULL
  gen_agent_hourly_total_response_tracker <- merge(gen_agent_hourly_total_response_tracker, gen_total_response_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("gen_team_hourly_total_response_tracker"))
{
  gen_team_hourly_total_response_tracker <- gen_total_response_per_team
} else {
  gen_team_hourly_total_response_tracker$total_response_target <- NULL
  gen_team_hourly_total_response_tracker$no_of_agents <- NULL
  gen_team_hourly_total_response_tracker <- merge(gen_team_hourly_total_response_tracker, gen_total_response_per_team, by = "tl_name", all = T)
}

write.csv(gen_agent_hourly_total_response_tracker, file = paste0("gen_agent_hourly_total_response_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_team_hourly_total_response_tracker, file = paste0("gen_team_hourly_total_response_tracker",hour(Sys.time()),".csv"),row.names = F)

## valid Response
gen_valid_response_per_agent <- gen_responses[,c("agent_id", "Valid.response")]

gen_valid_response_per_team <- merge(gen_valid_response_per_agent, gen_survey_agents[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
gen_valid_response_per_team <- gen_valid_response_per_team %>% group_by(tl_name) %>% summarise(valid_response = sum(Valid.response))

gen_valid_response_per_agent$valid_response_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*3, (hour(Sys.time()) - 10)*3)

gen_valid_response_per_team <- merge(gen_valid_response_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
gen_valid_response_per_team$valid_response_target <- gen_total_response_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*3, (hour(Sys.time()) - 10)*3)

names(gen_valid_response_per_agent)[2] <- paste0("valid_response_",hour(Sys.time()))
names(gen_valid_response_per_team)[2] <- paste0("valid_response_",hour(Sys.time()))

## agentwise
if(!exists("gen_agent_hourly_valid_response_tracker"))
{
  gen_agent_hourly_valid_response_tracker <- gen_valid_response_per_agent
} else {
  gen_agent_hourly_valid_response_tracker$valid_response_target <- NULL
  gen_agent_hourly_valid_response_tracker <- merge(gen_agent_hourly_valid_response_tracker, gen_valid_response_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("gen_team_hourly_valid_response_tracker"))
{
  gen_team_hourly_valid_response_tracker <- gen_valid_response_per_team
} else {
  gen_team_hourly_valid_response_tracker$valid_response_target <- NULL
  gen_team_hourly_valid_response_tracker$no_of_agents <- NULL
  gen_team_hourly_valid_response_tracker <- merge(gen_team_hourly_valid_response_tracker, gen_valid_response_per_team, by = "tl_name", all = T)
}

write.csv(gen_agent_hourly_valid_response_tracker, file = paste0("gen_agent_hourly_valid_response_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_team_hourly_valid_response_tracker, file = paste0("gen_team_hourly_valid_response_tracker",hour(Sys.time()),".csv"),row.names = F)

#################   ACCURACY 
gen_ques_accuracy_agent <- gen_survey_qc[,c("agent_id", "question_level_accuracy")]
gen_call_accuracy_agent <- gen_survey_qc[,c("agent_id", "call_level_accuracy")]

gen_ques_accuracy_team <- merge(gen_ques_accuracy_agent, attendance[,c("agent_id", "tl_name")], by= "agent_id",all.x = T)
gen_ques_accuracy_team <- gen_ques_accuracy_team %>% group_by(tl_name) %>% summarise(ques_accuracy = mean(question_level_accuracy))

gen_call_accuracy_team <- merge(gen_call_accuracy_agent, attendance[,c("agent_id", "tl_name")], by= "agent_id",all.x = T)
gen_call_accuracy_team <- gen_call_accuracy_team %>% group_by(tl_name) %>% summarise(call_accuracy = mean(call_level_accuracy))

names(gen_ques_accuracy_agent)[2] <- paste0("ques_accuracy_",hour(Sys.time()))
names(gen_call_accuracy_agent)[2] <- paste0("call_accuracy_",hour(Sys.time()))

names(gen_ques_accuracy_team)[2] <- paste0("ques_accuracy_",hour(Sys.time()))
names(gen_call_accuracy_team)[2] <- paste0("call_accuracy_",hour(Sys.time()))

## agentwise
if(!exists("gen_agent_hourly_ques_accuracy_tracker"))
{
  gen_agent_hourly_ques_accuracy_tracker <- gen_ques_accuracy_agent
} else {
  gen_agent_hourly_ques_accuracy_tracker <- merge(gen_agent_hourly_ques_accuracy_tracker, gen_ques_accuracy_agent, by = "agent_id", all = T)
}

if(!exists("gen_agent_hourly_call_accuracy_tracker"))
{
  gen_agent_hourly_call_accuracy_tracker <- gen_call_accuracy_agent
} else {
  gen_agent_hourly_call_accuracy_tracker <- merge(gen_agent_hourly_call_accuracy_tracker, gen_call_accuracy_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("gen_team_hourly_ques_accuracy_tracker"))
{
  gen_team_hourly_ques_accuracy_tracker <- gen_ques_accuracy_team
} else {
  gen_team_hourly_ques_accuracy_tracker <- merge(gen_team_hourly_ques_accuracy_tracker, gen_ques_accuracy_team, by = "tl_name", all = T)
}

if(!exists("gen_team_hourly_call_accuracy_tracker"))
{
  gen_team_hourly_call_accuracy_tracker <- gen_call_accuracy_team
} else {
  gen_team_hourly_call_accuracy_tracker <- merge(gen_team_hourly_call_accuracy_tracker, gen_call_accuracy_team, by = "tl_name", all = T)
}
write.csv(gen_agent_hourly_ques_accuracy_tracker, file = paste0("gen_agent_hourly_ques_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_agent_hourly_call_accuracy_tracker, file = paste0("gen_agent_hourly_call_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)

write.csv(gen_team_hourly_ques_accuracy_tracker, file = paste0("gen_team_hourly_ques_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_team_hourly_call_accuracy_tracker, file = paste0("gen_team_hourly_call_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)

######################################  GENERAL OVERALL SUMMARY  ########################

gen_overall <- merge(gen_survey_agents, gen_dialed_nums_per_agent, by = "agent_id", all.y = T)
gen_overall <- merge(gen_overall, gen_first_last[,c("agent_id", "First.Response", "Last.Response")], by = "agent_id", all.x = T)
gen_overall <- merge(gen_overall, gen_total_response_per_agent, by = "agent_id", all.x = T)
gen_overall <- merge(gen_overall, gen_valid_response_per_agent, by = "agent_id", all.x = T)
gen_overall <- merge(gen_overall, gen_survey_qc[c("agent_id", "calls_qc'd")], by = "agent_id", all.x = T)
gen_overall <- merge(gen_overall, gen_ques_accuracy_agent, by = "agent_id", all.x = T)
gen_overall <- merge(gen_overall, gen_call_accuracy_agent, by = "agent_id", all.x = T)
gen_overall$total_to_valid_conversion <- round(gen_overall[,paste0("valid_response_",hour(Sys.time()))]/gen_overall[,paste0("total_response_",hour(Sys.time()))]*100,2)
gen_overall <- gen_overall[order(gen_overall$tl_name, gen_overall$total_dialed, decreasing = T ),]
gen_overall$`calls_qc'd`[is.na(gen_overall$`calls_qc'd`)] <- 0

gen_overall$dialed_score <- gen_overall[,c(paste0("total_dialed_", hour(Sys.time())))] * 1 
gen_overall$submitted_forms_score <- gen_overall[,c(paste0("total_response_", hour(Sys.time())))] * 5
gen_overall$valid_forms_score <- gen_overall[,c(paste0("valid_response_", hour(Sys.time())))] * 20 
                                 
gen_overall$ques_acc_score <- ifelse(gen_overall$`calls_qc'd` >=2 & gen_overall[,c(paste0("ques_accuracy_", hour(Sys.time())))] >= 95, 100, 0)
gen_overall$call_acc_score <- ifelse(gen_overall$`calls_qc'd` >=2 & gen_overall[,c(paste0("call_accuracy_", hour(Sys.time())))] >= 75, 200, 0)

gen_overall$total_score <- sapply(X = 1:nrow(gen_overall),
                                  FUN = function(x)
                                    {
                                    sum(gen_overall[x,names(gen_overall)[names(gen_overall) %like% "score"]],na.rm = T)
                                  })

gen_overall_team <- gen_overall %>% group_by(tl_name) %>% summarise(dialed_score = round(mean(dialed_score, na.rm = T)),
                                                                    submitted_forms_score = round(mean(submitted_forms_score, na.rm = T)),
                                                                    valid_forms_score = round(mean(valid_forms_score, na.rm = T)),
                                                                    ques_acc_score = round(mean(ques_acc_score, na.rm = T)),
                                                                    call_acc_score = round(mean(call_acc_score, na.rm = T)),
                                                                    total_team_points = round(mean(total_score, na.rm = T)))

gen_overall_team <- gen_overall_team[!is.na(gen_overall_team$tl_name) & !(gen_overall_team$tl_name %like% "FRESHERS"),]
gen_overall_team <- gen_overall_team[order(gen_overall_team$total_team_points, decreasing = T),]

gen_wb <- createWorkbook(paste("General_Survey_Overall_Report",hour(Sys.time()),sep = "_")) 

addWorksheet(gen_wb, "Agentwise Overall Report")
writeData(gen_wb, "Agentwise Overall Report", gen_overall, rowNames = F)

addWorksheet(gen_wb, "Teamwise Overall Report")
writeData(gen_wb, "Teamwise Overall Report", gen_overall_team, rowNames = F)

saveWorkbook(gen_wb, paste0(paste("General_Survey_Overall_Report",hour(Sys.time()),sep = "_"),".xlsx"), overwrite = T)

#################     AB SURVEY TRACK   ###################################

ab_survey_qc <- list.files()[list.files() %like% "QC" & (list.files() %like% "AB" | list.files() %like% "Candidate")]
ab_survey_qc <- read_excel(ab_survey_qc, sheet = 2)
ab_survey_qc$X <- NULL
names(ab_survey_qc) <- tolower(names(ab_survey_qc))
ab_survey_qc$agent_id <- gsub("^\\s|\\s$","",ab_survey_qc$agent_id)

response_file <- list.files()[list.files() %like% "Summary"]
ab_hourly_total <- read_excel(response_file, sheet = 3)
ab_hourly_complete <- read_excel(response_file, sheet = 4)

ab_dispo <- disposition[disposition$CampName %like% "candidate",]
ab_dispo_distrib <- as.data.frame(table(ab_dispo$Disposition))
names(ab_dispo_distrib) <- c("disposition", "frequency")
ab_dispo_distrib$disposition <- as.character(ab_dispo_distrib$disposition)
ab_dispo_distrib$total_dialed <- nrow(ab_dispo_distrib)
ab_dispo_distrib$percent <- round(ab_dispo_distrib$frequency/ab_dispo_distrib$total_dialed*100,2)  

#################   NUMBER OF CALLS 
ab_dialed_nums_per_agent <- as.data.frame(ab_dispo %>% group_by(agent_id) %>% summarise(total_numbers_dialed = length(custno)))

ab_dialed_nums_per_team <- merge(ab_dialed_nums_per_agent, attendance[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_dialed_nums_per_team <- as.data.frame(ab_dialed_nums_per_team %>% group_by(tl_name) %>% summarise(total_numbers_dialed = sum(total_numbers_dialed)))

ab_dialed_nums_per_agent$dialed_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*25, (hour(Sys.time()) - 10)*25)

ab_dialed_nums_per_team <- merge(ab_dialed_nums_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
ab_dialed_nums_per_team$dialed_target <- ab_dialed_nums_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*25, (hour(Sys.time()) - 10)*25)

names(ab_dialed_nums_per_agent)[2] <- paste0("total_dialed_",hour(Sys.time()))
names(ab_dialed_nums_per_team)[2] <- paste0("total_dialed_",hour(Sys.time()))

## agentwise
if(!exists("ab_agent_hourly_dialed_num_tracker"))
{
  ab_agent_hourly_dialed_num_tracker <- ab_dialed_nums_per_agent
} else {
  ab_agent_hourly_dialed_num_tracker$dialed_target <- NULL
  ab_agent_hourly_dialed_num_tracker <- merge(ab_agent_hourly_dialed_num_tracker, ab_dialed_nums_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("ab_team_hourly_dialed_num_tracker"))
{
  ab_team_hourly_dialed_num_tracker <- ab_dialed_nums_per_team
} else {
  ab_team_hourly_dialed_num_tracker$dialed_target <- NULL
  ab_team_hourly_dialed_num_tracker$no_of_agents <- NULL
  ab_team_hourly_dialed_num_tracker <- merge(ab_team_hourly_dialed_num_tracker, gen_dialed_nums_per_team, by = "tl_name", all = T)
}

write.csv(ab_agent_hourly_dialed_num_tracker, file = paste0("ab_agent_hourly_dialed_num_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(ab_team_hourly_dialed_num_tracker, file = paste0("ab_team_hourly_dialed_num_tracker",hour(Sys.time()),".csv"),row.names = F)

#################   NUMBER OF FORMS

## Total Response
ab_hourly_total <- ab_hourly_total[,c("agent_id", "total_responses")]
ab_total_response_per_agent <- ab_hourly_total %>% group_by(agent_id) %>% summarise(Total.response = sum(total_responses))
ab_total_response_per_agent <- merge(attendance[,c("agent_id", "wfh_status")], ab_total_response_per_agent, by = "agent_id", all.y = T)

ab_total_response_per_team <- merge(ab_total_response_per_agent, attendance[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_total_response_per_team <- as.data.frame(ab_total_response_per_team %>% group_by(tl_name, wfh_status) %>% summarise(total_response = sum(Total.response)))

ab_total_response_per_agent$total_response_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*5, (hour(Sys.time()) - 10)*5)
ab_total_response_per_team <- merge(ab_total_response_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
ab_total_response_per_team$total_response_target <- ab_total_response_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*5, (hour(Sys.time()) - 10)*5)

names(ab_total_response_per_agent)[3] <- paste0("total_response_",hour(Sys.time()))
names(ab_total_response_per_team)[3] <- paste0("total_response_",hour(Sys.time()))

## agentwise
if(!exists("ab_agent_hourly_total_response_tracker"))
{
  ab_agent_hourly_total_response_tracker <- ab_total_response_per_agent
} else {
  ab_agent_hourly_total_response_tracker$wfh_status <- NULL
  ab_agent_hourly_total_response_tracker$total_response_target <- NULL
  ab_agent_hourly_total_response_tracker <- merge(ab_agent_hourly_total_response_tracker, ab_total_response_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("ab_team_hourly_total_response_tracker"))
{
  ab_team_hourly_total_response_tracker <- ab_total_response_per_team
} else {
  ab_team_hourly_total_response_tracker$wfh_status <- NULL
  ab_team_hourly_total_response_tracker$no_of_agents <- NULL
  ab_team_hourly_total_response_tracker$total_response_target <- NULL
  ab_team_hourly_total_response_tracker <- merge(ab_team_hourly_total_response_tracker, ab_total_response_per_team, by = "tl_name", all = T)
}

write.csv(ab_agent_hourly_total_response_tracker, file = paste0("ab_agent_hourly_total_response_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(ab_team_hourly_total_response_tracker, file = paste0("ab_team_hourly_total_response_tracker",hour(Sys.time()),".csv"),row.names = F)

## valid Response
ab_hourly_complete <- ab_hourly_complete[,c("agent_id", "total_valid")]
ab_valid_response_per_agent <- ab_hourly_complete %>% group_by(agent_id) %>% summarise(Valid.response = sum(total_valid))
ab_valid_response_per_agent <- merge(attendance[,c("agent_id", "wfh_status")], ab_valid_response_per_agent, by = "agent_id", all.y = T)

ab_valid_response_per_team <- merge(ab_valid_response_per_agent, attendance[,c("agent_id", "tl_name")], by = "agent_id", all.x = T)
ab_valid_response_per_team <- as.data.frame(ab_valid_response_per_team %>% group_by(tl_name) %>% summarise(valid_response = sum(Valid.response)))

ab_valid_response_per_agent$valid_response_target <- ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*4, (hour(Sys.time()) - 10)*4)
ab_valid_response_per_team <- merge(ab_valid_response_per_team, no_of_agents_per_TL, by = "tl_name", all.x = T)
ab_valid_response_per_team$valid_response_target <- ab_valid_response_per_team$no_of_agents*ifelse(hour(Sys.time()) <= 13, (hour(Sys.time()) - 9)*4, (hour(Sys.time()) - 10)*4)

names(ab_valid_response_per_agent)[3] <- paste0("valid_response_",hour(Sys.time()))
names(ab_valid_response_per_team)[2] <- paste0("valid_response_",hour(Sys.time()))

## agentwise
if(!exists("ab_agent_hourly_valid_response_tracker"))
{
  ab_agent_hourly_valid_response_tracker <- ab_valid_response_per_agent
} else {
  ab_agent_hourly_valid_response_tracker$wfh_status <- NULL
  ab_agent_hourly_valid_response_tracker$valid_response_target <- NULL
  ab_agent_hourly_valid_response_tracker <- merge(ab_agent_hourly_valid_response_tracker, ab_valid_response_per_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("ab_team_hourly_valid_response_tracker"))
{
  ab_team_hourly_valid_response_tracker <- ab_valid_response_per_team
} else{
  ab_team_hourly_valid_response_tracker$no_of_agents <- NULL
  ab_team_hourly_valid_response_tracker$valid_response_target <- NULL
  ab_team_hourly_valid_response_tracker <- merge(ab_team_hourly_valid_response_tracker, ab_valid_response_per_team, by = "tl_name", all = T)
}

write.csv(ab_agent_hourly_valid_response_tracker, file = paste0("ab_agent_hourly_valid_response_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(ab_team_hourly_valid_response_tracker, file = paste0("ab_team_hourly_valid_response_tracker",hour(Sys.time()),".csv"),row.names = F)

#################   ACCURACY 
ab_ques_accuracy_agent <- ab_survey_qc[,c("agent_id", "question_level_accuracy")]
ab_call_accuracy_agent <- ab_survey_qc[,c("agent_id", "call_level_accuracy")]

ab_ques_accuracy_team <- merge(ab_ques_accuracy_agent, attendance[,c("agent_id", "tl_name")], by= "agent_id",all.x = T)
ab_ques_accuracy_team <- ab_ques_accuracy_team %>% group_by(tl_name) %>% summarise(ques_accuracy = mean(question_level_accuracy))

ab_call_accuracy_team <- merge(ab_call_accuracy_agent, attendance[,c("agent_id", "tl_name")], by= "agent_id",all.x = T)
ab_call_accuracy_team <- ab_call_accuracy_team %>% group_by(tl_name) %>% summarise(call_accuracy = mean(call_level_accuracy))

names(ab_ques_accuracy_agent)[2] <- paste0("ques_accuracy_",hour(Sys.time()))
names(ab_call_accuracy_agent)[2] <- paste0("call_accuracy_",hour(Sys.time()))

names(ab_ques_accuracy_team)[2] <- paste0("ques_accuracy_",hour(Sys.time()))
names(ab_call_accuracy_team)[2] <- paste0("call_accuracy_",hour(Sys.time()))

## agentwise
if(!exists("ab_agent_hourly_ques_accuracy_tracker"))
{
  ab_agent_hourly_ques_accuracy_tracker <- ab_ques_accuracy_agent
} else {
  ab_agent_hourly_ques_accuracy_tracker <- merge(ab_agent_hourly_ques_accuracy_tracker, ab_ques_accuracy_agent, by = "agent_id", all = T)
}

if(!exists("ab_agent_hourly_call_accuracy_tracker"))
{
  ab_agent_hourly_call_accuracy_tracker <- ab_call_accuracy_agent
} else {
  ab_agent_hourly_call_accuracy_tracker <- merge(ab_agent_hourly_call_accuracy_tracker, ab_call_accuracy_agent, by = "agent_id", all = T)
}

## teamwise
if(!exists("ab_team_hourly_ques_accuracy_tracker"))
{
  ab_team_hourly_ques_accuracy_tracker <- ab_ques_accuracy_team
} else {
  ab_team_hourly_ques_accuracy_tracker <- merge(ab_team_hourly_ques_accuracy_tracker, ab_ques_accuracy_team, by = "tl_name", all = T)
}

if(!exists("ab_team_hourly_call_accuracy_tracker"))
{
  ab_team_hourly_call_accuracy_tracker <- ab_call_accuracy_team
} else {
  ab_team_hourly_call_accuracy_tracker <- merge(ab_team_hourly_call_accuracy_tracker, ab_call_accuracy_team, by = "tl_name", all = T)
}
write.csv(ab_agent_hourly_ques_accuracy_tracker, file = paste0("ab_agent_hourly_ques_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(ab_agent_hourly_call_accuracy_tracker, file = paste0("ab_agent_hourly_call_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)

write.csv(gen_team_hourly_ques_accuracy_tracker, file = paste0("gen_team_hourly_ques_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)
write.csv(gen_team_hourly_call_accuracy_tracker, file = paste0("gen_team_hourly_call_accuracy_tracker",hour(Sys.time()),".csv"),row.names = F)


######################################  AB OVERALL SUMMARY  ########################

ab_overall <- merge(ab_survey_agents, ab_dialed_nums_per_agent, by = "agent_id", all = T)
#ab_overall <- merge(ab_overall, ab_first_last[,c("agent_id", "First.Response", "Last.Response")], by = "agent_id", all = T)
ab_overall <- merge(ab_overall, ab_total_response_per_agent, by = c("agent_id","wfh_status"), all.x = T)
ab_overall <- merge(ab_overall, ab_valid_response_per_agent, by = c("agent_id", "wfh_status"), all.x = T)
ab_overall <- merge(ab_overall, ab_survey_qc[,c("agent_id", "calls_qc'd")], by = "agent_id", all.x = T)
ab_overall <- merge(ab_overall, ab_ques_accuracy_agent, by = "agent_id", all.x = T)
ab_overall <- merge(ab_overall, ab_call_accuracy_agent, by = "agent_id", all.x = T)
ab_overall$total_to_valid_conversion <- round(ab_overall[,paste0("valid_response_",hour(Sys.time()))]/ab_overall[,paste0("total_response_",hour(Sys.time()))]*100,2)
ab_overall$`calls_qc'd`[is.na(ab_overall$`calls_qc'd`)] <- 0
ab_overall <- ab_overall[order(ab_overall$tl_name, ab_overall$total_dialed, decreasing = T ),]

ab_overall$dialed_score <- ab_overall[,c(paste0("total_dialed_", hour(Sys.time())))] * 1 
ab_overall$submitted_forms_score <- ab_overall[,c(paste0("total_response_", hour(Sys.time())))] * 5
ab_overall$valid_forms_score <- ab_overall[,c(paste0("valid_response_", hour(Sys.time())))] * 20 

ab_overall$ques_acc_score <- ifelse(ab_overall$`calls_qc'd` > 2 & ab_overall[,(paste0("ques_accuracy_", hour(Sys.time())))] >= 100, 100, 0)
ab_overall$call_acc_score <- ifelse(ab_overall$`calls_qc'd` > 2 & ab_overall[,(paste0("call_accuracy_", hour(Sys.time())))] >= 95, 200, 0)

ab_overall$total_score <- sapply(X = 1:nrow(ab_overall),
                                  FUN = function(x)
                                  {
                                    sum(ab_overall[x,names(ab_overall)[names(ab_overall) %like% "score"]],na.rm = T)
                                  })

ab_overall_team <- ab_overall %>% group_by(tl_name) %>% summarise(dialed_score = round(mean(dialed_score, na.rm = T)),
                                                                    submitted_forms_score = round(mean(submitted_forms_score, na.rm = T)),
                                                                    valid_forms_score = round(mean(valid_forms_score, na.rm = T)),
                                                                    ques_acc_score = round(mean(ques_acc_score, na.rm = T)),
                                                                    call_acc_score = round(mean(call_acc_score, na.rm = T)),
                                                                    total_team_points = round(mean(total_score, na.rm = T)))

ab_overall_team <- ab_overall_team[!is.na(ab_overall_team$tl_name) & !(ab_overall_team$tl_name %like% "FRESHERS"),]
ab_overall_team <- ab_overall_team[order(ab_overall_team$total_team_points, decreasing = T),]

ab_wb <- createWorkbook(paste("Candidate_Survey_Overall_Report",hour(Sys.time()),sep = "_")) 

addWorksheet(ab_wb, "Agentwise Overall Report")
writeData(ab_wb, "Agentwise Overall Report", ab_overall, rowNames = F)

addWorksheet(ab_wb, "Teamwise Overall Report")
writeData(ab_wb, "Teamwise Overall Report", ab_overall_team, rowNames = F)

saveWorkbook(ab_wb, paste0(paste("Candidate_Survey_Overall_Report",hour(Sys.time()),sep = "_"),".xlsx"), overwrite = T)

################################   ERROR TRACKING  ##########################################

gen_report <- list.files()[list.files() %like% "General Survey QC"]
cand_report <- list.files()[list.files() %like% "Candidate Survey QC"]
ab_report <- list.files()[list.files() %like% "AB candidate QC"]

############ General Survey Errors
gen_errors <- read_excel(gen_report, sheet = 1)
gen_errors <- merge(attendance[,c("agent_id", "agent_name", "tl_name")], gen_errors, by = "agent_id")
gen_errors <- cbind(time = rep(hour(Sys.time()),nrow(gen_errors)),gen_errors)

if(!exists("gen_agent_hourly_error_tracker"))
{
  gen_agent_hourly_error_tracker <- gen_errors
  gen_agent_hourly_error_tracker <- gen_agent_hourly_error_tracker[order(gen_agent_hourly_error_tracker$tl_name,gen_agent_hourly_error_tracker$agent_id, gen_agent_hourly_error_tracker$time, decreasing = F),]
} else {
  gen_errors <- gen_errors[!gen_errors$phone_no %in% gen_agent_hourly_error_tracker$phone_no,]
  gen_agent_hourly_error_tracker <- rbind(gen_agent_hourly_error_tracker,gen_errors)
  gen_agent_hourly_error_tracker <- gen_agent_hourly_error_tracker[order(gen_agent_hourly_error_tracker$tl_name,gen_agent_hourly_error_tracker$agent_id, gen_agent_hourly_error_tracker$time, decreasing = F),]
}
write.csv(gen_agent_hourly_error_tracker, file = "gen_agent_hourly_error_tracker.csv", row.names = F)

############ Candidate Survey Errors
cand_errors <- read_excel(cand_report, sheet = 1)
cand_errors <- merge(attendance[,c("agent_id", "agent_name", "tl_name")], cand_errors, by = "agent_id")
cand_errors <- cbind(time = rep(hour(Sys.time()),nrow(cand_errors)),cand_errors)

if(!exists("cand_agent_hourly_error_tracker"))
{
  cand_agent_hourly_error_tracker <- cand_errors
  cand_agent_hourly_error_tracker <- cand_agent_hourly_error_tracker[order(cand_agent_hourly_error_tracker$tl_name,cand_agent_hourly_error_tracker$agent_id, cand_agent_hourly_error_tracker$time, decreasing = F),]
} else {
  cand_errors <- cand_errors[!cand_errors$phone_no %in% cand_agent_hourly_error_tracker$phone_no,]
  cand_agent_hourly_error_tracker <- rbind(cand_agent_hourly_error_tracker,cand_errors)
  cand_agent_hourly_error_tracker <- cand_agent_hourly_error_tracker[order(cand_agent_hourly_error_tracker$tl_name,cand_agent_hourly_error_tracker$agent_id, cand_agent_hourly_error_tracker$time, decreasing = F),]
}
write.csv(cand_agent_hourly_error_tracker, file = "cand_agent_hourly_error_tracker.csv", row.names = F)

############ AB Survey Errors
ab_errors <- read_excel(ab_report, sheet = 1)
ab_errors <- merge(attendance[,c("agent_id", "agent_name", "tl_name")], ab_errors, by = "agent_id")
ab_errors <- cbind(time = rep(hour(Sys.time()),nrow(ab_errors)),ab_errors)

if(!exists("ab_agent_hourly_error_tracker"))
{
  ab_agent_hourly_error_tracker <- ab_errors
  ab_agent_hourly_error_tracker <- ab_agent_hourly_error_tracker[order(ab_agent_hourly_error_tracker$tl_name,ab_agent_hourly_error_tracker$agent_id, ab_agent_hourly_error_tracker$time, decreasing = F),]
} else {
  ab_errors <- ab_errors[!ab_errors$phone_no %in% ab_agent_hourly_error_tracker$phone_no,]
  ab_agent_hourly_error_tracker <- rbind(ab_agent_hourly_error_tracker,ab_errors)
  ab_agent_hourly_error_tracker <- ab_agent_hourly_error_tracker[order(ab_agent_hourly_error_tracker$tl_name,ab_agent_hourly_error_tracker$agent_id, ab_agent_hourly_error_tracker$time, decreasing = F),]
}
write.csv(ab_agent_hourly_error_tracker, file = "ab_agent_hourly_error_tracker.csv", row.names = F)

