setwd("F:/IPAC/Chennai Telephony/daily_and_hourly_trackers/daily trackers/error_tracking/")
track_date <- "2021-03-01"
#################   CREATE DATE FOLDER
# if(dir.exists(as.character(Sys.Date())))
# {
#   message("Folder exists")
# } else {
#   dir.create(as.character(Sys.Date()))
#   message("Folder created")
# }

files <- list.files()[grepl("csv|xlsx",list.files())]
gen_report <- list.files()[list.files() %like% "General Survey QC"]
cand_report <- list.files()[list.files() %like% "Candidate Survey QC"]
ab_report <- list.files()[list.files() %like% "AB candidate QC"]
error_att_file <- list.files()[list.files() %like% "Attend"]

# move_file <- function(file_name,from,to){
#   file.rename( from = file.path(paste(from, file_name, sep ="/")) ,
#                to = file.path(paste(to,file_name,sep ="/")))
# }
# 
# #################   MOVE FILES TO DATE FOLDER
# sapply(files, move_file, from = getwd(), to = paste(getwd(),Sys.Date(),sep = "/"))
# setwd(paste(getwd(),Sys.Date(),sep ="/"))

# error_att <- attendance
error_att <- read_excel(error_att_file, sheet = 1)
error_att <- error_att[,1:6]
names(error_att) <- c("s_no", "agent_id", "tl_name", "agent_name", "wfh_status", "calling_type")
error_att$s_no <- NULL
error_att$agent_id <- as.character(error_att$agent_id)
error_att$calling_type <- tolower(error_att$calling_type)
error_att$tl_name <- toupper(error_att$tl_name)
error_att$wfh_status <- toupper(error_att$wfh_status)
error_att <- error_att[!(is.na(error_att$agent_id)),]

gen_error_agents <- error_att[(!error_att$tl_name %like% "TRAINING")&
                                (error_att$calling_type %in% c("tn survey", "general survey", "genral survey recalling", "survey re-calling")),]
ab_error_agents <- error_att[ error_att$calling_type %like% "ab testing" | 
                                  error_att$calling_type %like% "a/b testing",] 
cand_error_agents <- error_att[error_att$calling_type %like% "candidate survey" | 
                                  error_att$calling_type %like% "candidate general",]

error_att <- rbind(gen_error_agents, cand_error_agents, ab_error_agents)

############ General Survey Errors
gen_errors <- read_excel(gen_report, sheet = 1)
gen_errors <- merge(error_att[,c("agent_id", "agent_name", "tl_name")], gen_errors, by = "agent_id")
gen_errors$agent_name <- toupper(gen_errors$agent_name)
gen_errors <- cbind(date = as.character(rep(track_date,nrow(gen_errors))),gen_errors)
# gen_errors <- cbind(date = rep(Sys.Date(),nrow(gen_errors)),gen_errors)
gen_errors$date <- as.POSIXct(as.character(gen_errors$date), format = "%Y-%m-%d")


gen_errors_agg <- read.csv("gen_errors_agg.csv", stringsAsFactors = F)
gen_errors_agg$agent_name <- toupper(gen_errors_agg$agent_name)
gen_errors_agg$date <- as.POSIXct(gen_errors_agg$date, format = "%Y-%m-%d")
gen_errors_agg <- rbind(gen_errors_agg,gen_errors)
write.csv(gen_errors_agg, file = "gen_errors_agg.csv", row.names = F)

############ AB Survey Errors
ab_errors <- read_excel(ab_report, sheet = 1)
ab_errors <- merge(error_att[,c("agent_id", "agent_name", "tl_name")], ab_errors, by = "agent_id")
ab_errors$agent_name <- toupper(ab_errors$agent_name)
ab_errors <- cbind(date = as.character(rep(track_date,nrow(ab_errors))),ab_errors)
# ab_errors <- cbind(date = rep(Sys.Date(),nrow(ab_errors)),ab_errors)
ab_errors$date <- as.POSIXct(as.character(ab_errors$date), format = "%Y-%m-%d")

ab_errors_agg <- read.csv("ab_errors_agg.csv", stringsAsFactors = F)
ab_errors_agg$agent_name <- toupper(ab_errors_agg$agent_name)
ab_errors_agg$date <- as.POSIXct(ab_errors_agg$date, format = "%Y-%m-%d")
ab_errors_agg <- rbind(ab_errors_agg,ab_errors)
write.csv(ab_errors_agg, file = "ab_errors_agg.csv", row.names = F)

############ Candidate Survey Errors
cand_errors <- read_excel(cand_report, sheet = 1)
cand_errors <- merge(error_att[,c("agent_id", "agent_name", "tl_name")], cand_errors, by = "agent_id")
cand_errors$agent_name <- toupper(cand_errors$agent_name)
# cand_errors <- cbind(date = rep(Sys.Date(),nrow(cand_errors)),cand_errors)
cand_errors <- cbind(date = as.character(rep(track_date,nrow(cand_errors))),cand_errors)
cand_errors$date <- as.POSIXct(as.character(cand_errors$date), format = "%Y-%m-%d")

cand_errors_agg <- read.csv("cand_errors_agg.csv", stringsAsFactors = F)
cand_errors_agg$agent_name <- toupper(cand_errors_agg$agent_name)
cand_errors_agg$date <- as.POSIXct(cand_errors_agg$date, format = "%Y-%m-%d")
cand_errors_agg <- rbind(cand_errors_agg,cand_errors)
write.csv(cand_errors_agg, file = "cand_errors_agg.csv", row.names = F)

##############################   ANALYSIS   ##########################################

#################### General Survey #############################
gen_errors_agentwise <- as.data.frame(gen_errors %>% group_by(date,agent_id,agent_name,tl_name) %>%
  summarise(district = sum(district == 0),
            ac = sum(ac == 0),
            gender = sum(gender == 0),
            age = sum(age == 0),
            living = sum(living == 0),
            village_panchayat = sum(village_panchayat == 0),
            block = sum(block == 0),
            municipality = sum(muncipality == 0),
            party_vote_2019_election = sum(party_vote_2019_election == 0),
            party_vote_n = sum(party_vote_n == 0),
            vote_party_now = sum(vote_party_now == 0),
            see_chief_minister = sum(see_chief_minister == 0),
            rating_tn_govt = sum(rating_tn_govt == 0),
            social_category = sum(social_category == 0),
            religion = sum(religion == 0),
            caste = sum(caste == 0),
            occupation = sum(occupation == 0)))
gen_errors_agentwise$total_errors <- rowSums(gen_errors_agentwise[5:length(gen_errors_agentwise)])
gen_errors_agentwise <- gen_errors_agentwise[order(gen_errors_agentwise$tl_name,gen_errors_agentwise$agent_id, gen_errors_agentwise$date, decreasing = F),]

gen_errors_agentwise_agg <- as.data.frame(gen_errors_agg %>% group_by(agent_id) %>%
                                        summarise(no_of_days = length(unique(date)),
                                                  district = sum(district == 0),
                                                  ac = sum(ac == 0),
                                                  gender = sum(gender == 0),
                                                  age = sum(age == 0),
                                                  living = sum(living == 0),
                                                  village_panchayat = sum(village_panchayat == 0),
                                                  block = sum(block == 0),
                                                  municipality = sum(muncipality == 0),
                                                  party_vote_2019_election = sum(party_vote_2019_election == 0),
                                                  party_vote_n = sum(party_vote_n == 0),
                                                  vote_party_now = sum(vote_party_now == 0),
                                                  see_chief_minister = sum(see_chief_minister == 0),
                                                  rating_tn_govt = sum(rating_tn_govt == 0),
                                                  social_category = sum(social_category == 0),
                                                  religion = sum(religion == 0),
                                                  caste = sum(caste == 0),
                                                  occupation = sum(occupation == 0)))
gen_errors_agentwise_agg$agg_total_errors <- rowSums(gen_errors_agentwise_agg[3:length(gen_errors_agentwise_agg)])
gen_errors_agentwise_agg$errors_per_day <- round(gen_errors_agentwise_agg$agg_total_errors/gen_errors_agentwise_agg$no_of_days,2) 
gen_errors_agentwise_agg <- unique(merge(gen_errors_agg[,c("agent_id","agent_name","tl_name")], gen_errors_agentwise_agg, by = "agent_id", all.y = T))
gen_errors_agentwise_agg <- as.data.frame(gen_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
gen_errors_agentwise_agg$rank <- NULL
gen_errors_agentwise_agg <- gen_errors_agentwise_agg[order(gen_errors_agentwise_agg$tl_name, gen_errors_agentwise_agg$errors_per_day, decreasing = T),]

gen_errors_teamwise <- as.data.frame(gen_errors %>% group_by(date,tl_name) %>%
                        summarise(district = sum(district == 0),
                                  ac = sum(ac == 0),
                                  gender = sum(gender == 0),
                                  age = sum(age == 0),
                                  living = sum(living == 0),
                                  village_panchayat = sum(village_panchayat == 0),
                                  block = sum(block == 0),
                                  municipality = sum(muncipality == 0),
                                  party_vote_2019_election = sum(party_vote_2019_election == 0),
                                  party_vote_n = sum(party_vote_n == 0),
                                  vote_party_now = sum(vote_party_now == 0),
                                  see_chief_minister = sum(see_chief_minister == 0),
                                  rating_tn_govt = sum(rating_tn_govt == 0),
                                  social_category = sum(social_category == 0),
                                  religion = sum(religion == 0),
                                  caste = sum(caste == 0),
                                  occupation = sum(occupation == 0)))
gen_errors_teamwise$total_errors <- rowSums(gen_errors_teamwise[3:length(gen_errors_teamwise)])
gen_errors_teamwise <- gen_errors_teamwise[order(gen_errors_teamwise$tl_name,gen_errors_teamwise$date, decreasing = F),]

gen_errors_teamwise_agg <- as.data.frame(gen_errors_agg %>% group_by(tl_name) %>%
                                            summarise(no_of_days = length(unique(date)),
                                                      district = sum(district == 0),
                                                      ac = sum(ac == 0),
                                                      gender = sum(gender == 0),
                                                      age = sum(age == 0),
                                                      living = sum(living == 0),
                                                      village_panchayat = sum(village_panchayat == 0),
                                                      block = sum(block == 0),
                                                      municipality = sum(muncipality == 0),
                                                      party_vote_2019_election = sum(party_vote_2019_election == 0),
                                                      party_vote_n = sum(party_vote_n == 0),
                                                      vote_party_now = sum(vote_party_now == 0),
                                                      see_chief_minister = sum(see_chief_minister == 0),
                                                      rating_tn_govt = sum(rating_tn_govt == 0),
                                                      social_category = sum(social_category == 0),
                                                      religion = sum(religion == 0),
                                                      caste = sum(caste == 0),
                                                      occupation = sum(occupation == 0)))
gen_errors_teamwise_agg$agg_total_errors <- rowSums(gen_errors_teamwise_agg[3:length(gen_errors_teamwise_agg)])
gen_errors_teamwise_agg$errors_per_day <- round(gen_errors_teamwise_agg$agg_total_errors/gen_errors_teamwise_agg$no_of_days,2) 
# gen_errors_teamwise_agg <- unique(merge(gen_errors_agg[,c("agent_id","agent_name","tl_name")], gen_errors_agentwise_agg, by = "agent_id", all.y = T))
# gen_errors_teamwise_agg <- as.data.frame(gen_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
# gen_errors_agentwise_agg$rank <- NULL
gen_errors_teamwise_agg <- gen_errors_teamwise_agg[order(gen_errors_teamwise_agg$errors_per_day, decreasing = T),]


gen_error_wb <- createWorkbook("gen_error_tracker")
daily_gen_errors_agentwise <- read_excel("gen_error_tracker.xlsx", sheet = 1)
daily_gen_errors_agentwise$date <- as.character(as.POSIXct(daily_gen_errors_agentwise$date, format = "%Y-%m-%d"))
gen_errors_agentwise$date <- as.character(gen_errors_agentwise$date)
daily_gen_errors_agentwise <- rbind(daily_gen_errors_agentwise,gen_errors_agentwise)
#daily_gen_errors_agentwise <- gen_errors_agentwise
addWorksheet(gen_error_wb, sheetName = "agent_level")
writeData(gen_error_wb, "agent_level", daily_gen_errors_agentwise, rowNames = F)

daily_gen_errors_teamwise <- read_excel("gen_error_tracker.xlsx", sheet = 2)
daily_gen_errors_teamwise$date <- as.character(as.POSIXct(daily_gen_errors_teamwise$date, format = "%Y-%m-%d"))
gen_errors_teamwise$date <- as.character(gen_errors_teamwise$date)
daily_gen_errors_teamwise <- rbind(daily_gen_errors_teamwise,gen_errors_teamwise)
#daily_gen_errors_teamwise <- gen_errors_teamwise
addWorksheet(gen_error_wb, sheetName = "team_level")
writeData(gen_error_wb, "team_level", daily_gen_errors_teamwise, rowNames = F)

addWorksheet(gen_error_wb, sheetName = "agent_level_aggregate")
writeData(gen_error_wb, "agent_level_aggregate", gen_errors_agentwise_agg, rowNames = F)

addWorksheet(gen_error_wb, sheetName = "team_level_aggregate")
writeData(gen_error_wb, "team_level_aggregate", gen_errors_teamwise_agg, rowNames = F)

saveWorkbook(gen_error_wb, file = "gen_error_tracker.xlsx", overwrite = T)

#################### AB Survey  ##################################
ab_errors_agentwise <- as.data.frame(ab_errors %>% group_by(date,agent_id,agent_name,tl_name) %>%
  summarise(district = sum(district == 0),
            ac = sum(ac == 0),
            gender = sum(gender == 0),
            age = sum(age == 0),
            party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
            party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
ab_errors_agentwise$total_errors <- rowSums(ab_errors_agentwise[5:length(ab_errors_agentwise)])
ab_errors_agentwise <- ab_errors_agentwise[order(ab_errors_agentwise$tl_name,ab_errors_agentwise$agent_id, ab_errors_agentwise$date, decreasing = F),]

ab_errors_agentwise_agg <- as.data.frame(ab_errors_agg %>% group_by(agent_id) %>%
                                       summarise(no_of_days = length(date),
                                                 district = sum(district == 0),
                                                 ac = sum(ac == 0),
                                                 gender = sum(gender == 0),
                                                 age = sum(age == 0),
                                                 party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                 party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
ab_errors_agentwise_agg$agg_total_errors <- rowSums(ab_errors_agentwise_agg[3:length(ab_errors_agentwise_agg)])
ab_errors_agentwise_agg$errors_per_day <- round(ab_errors_agentwise_agg$agg_total_errors/ab_errors_agentwise_agg$no_of_days,2) 
ab_errors_agentwise_agg <- unique(merge(ab_errors_agg[,c("agent_id","agent_name","tl_name")], ab_errors_agentwise_agg, by = "agent_id", all.y = T))
ab_errors_agentwise_agg <- as.data.frame(ab_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
ab_errors_agentwise_agg$rank <- NULL
ab_errors_agentwise_agg <- ab_errors_agentwise_agg[order(ab_errors_agentwise_agg$tl_name, ab_errors_agentwise_agg$errors_per_day, decreasing = T),]

ab_errors_teamwise <- as.data.frame(ab_errors %>% group_by(date,tl_name) %>%
                                       summarise(district = sum(district == 0),
                                                 ac = sum(ac == 0),
                                                 gender = sum(gender == 0),
                                                 age = sum(age == 0),
                                                 party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                 party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
ab_errors_teamwise$total_errors <- rowSums(ab_errors_teamwise[3:length(ab_errors_teamwise)])
ab_errors_teamwise <- ab_errors_teamwise[order(ab_errors_teamwise$tl_name, ab_errors_teamwise$date, decreasing = F),]

ab_errors_teamwise_agg <- as.data.frame(ab_errors_agg %>% group_by(tl_name) %>%
                                           summarise(no_of_days = length(unique(date)),
                                                     district = sum(district == 0),
                                                     ac = sum(ac == 0),
                                                     gender = sum(gender == 0),
                                                     age = sum(age == 0),
                                                     party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                     party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))

ab_errors_teamwise_agg$agg_total_errors <- rowSums(ab_errors_teamwise_agg[3:length(ab_errors_teamwise_agg)])
ab_errors_teamwise_agg$errors_per_day <- round(ab_errors_teamwise_agg$agg_total_errors/ab_errors_teamwise_agg$no_of_days,2) 
# gen_errors_teamwise_agg <- unique(merge(gen_errors_agg[,c("agent_id","agent_name","tl_name")], gen_errors_agentwise_agg, by = "agent_id", all.y = T))
# gen_errors_teamwise_agg <- as.data.frame(gen_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
# gen_errors_agentwise_agg$rank <- NULL
ab_errors_teamwise_agg <- ab_errors_teamwise_agg[order(ab_errors_teamwise_agg$errors_per_day, decreasing = T),]

ab_error_wb <- createWorkbook("ab_error_tracker")
daily_ab_errors_agentwise <- read_excel("ab_error_tracker.xlsx", sheet = 1)
daily_ab_errors_agentwise$date <- as.character(as.POSIXct(daily_ab_errors_agentwise$date, format = "%Y-%m-%d"))
ab_errors_agentwise$date <- as.character(ab_errors_agentwise$date)
daily_ab_errors_agentwise <- rbind(daily_ab_errors_agentwise,ab_errors_agentwise)
# daily_ab_errors_agentwise <- ab_errors_agentwise
addWorksheet(ab_error_wb, sheetName = "agent_level")
writeData(ab_error_wb, "agent_level", daily_ab_errors_agentwise, rowNames = F)

daily_ab_errors_teamwise <- read_excel("ab_error_tracker.xlsx", sheet = 2)
daily_ab_errors_teamwise$date <- as.character(as.POSIXct(daily_ab_errors_teamwise$date, format = "%Y-%m-%d"))
ab_errors_teamwise$date <- as.character(ab_errors_teamwise$date)
daily_ab_errors_teamwise <- rbind(daily_ab_errors_teamwise,ab_errors_teamwise)
# daily_ab_errors_teamwise <- ab_errors_teamwise
addWorksheet(ab_error_wb, sheetName = "team_level")
writeData(ab_error_wb, "team_level", daily_ab_errors_teamwise, rowNames = F)

addWorksheet(ab_error_wb, sheetName = "agent_level_aggregate")
writeData(ab_error_wb, "agent_level_aggregate", ab_errors_agentwise_agg, rowNames = F)

addWorksheet(ab_error_wb, sheetName = "team_level_aggregate")
writeData(ab_error_wb, "team_level_aggregate", ab_errors_teamwise_agg, rowNames = F)

saveWorkbook(ab_error_wb, file = "ab_error_tracker.xlsx", overwrite = T)

#################### Candidate Survey #######################
cand_errors_agentwise <- as.data.frame(cand_errors %>% group_by(date,agent_id,agent_name,tl_name) %>%
                                       summarise(district = sum(district == 0),
                                                 ac = sum(ac == 0),
                                                 gender = sum(gender == 0),
                                                 age = sum(age == 0),
                                                 party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                 party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
cand_errors_agentwise$total_errors <- rowSums(cand_errors_agentwise[5:length(cand_errors_agentwise)])
cand_errors_agentwise <- cand_errors_agentwise[order(cand_errors_agentwise$tl_name,cand_errors_agentwise$agent_id, cand_errors_agentwise$date, decreasing = F),]

cand_errors_agentwise_agg <- as.data.frame(cand_errors_agg %>% group_by(agent_id) %>%
                                           summarise(no_of_days = length(date),
                                                     district = sum(district == 0),
                                                     ac = sum(ac == 0),
                                                     gender = sum(gender == 0),
                                                     age = sum(age == 0),
                                                     party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                     party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
cand_errors_agentwise_agg$agg_total_errors <- rowSums(cand_errors_agentwise_agg[3:length(cand_errors_agentwise_agg)])
cand_errors_agentwise_agg$errors_per_day <- round(cand_errors_agentwise_agg$agg_total_errors/cand_errors_agentwise_agg$no_of_days,2) 
cand_errors_agentwise_agg <- unique(merge(cand_errors_agg[,c("agent_id","agent_name","tl_name")], cand_errors_agentwise_agg, by = "agent_id", all.y = T))
cand_errors_agentwise_agg <- as.data.frame(cand_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
cand_errors_agentwise_agg$rank <- NULL
cand_errors_agentwise_agg <- cand_errors_agentwise_agg[order(cand_errors_agentwise_agg$tl_name, cand_errors_agentwise_agg$errors_per_day, decreasing = T),]

cand_errors_teamwise <- as.data.frame(cand_errors %>% group_by(date,tl_name) %>%
                                      summarise(district = sum(district == 0),
                                                ac = sum(ac == 0),
                                                gender = sum(gender == 0),
                                                age = sum(age == 0),
                                                party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))
cand_errors_teamwise$total_errors <- rowSums(cand_errors_teamwise[3:length(cand_errors_teamwise)])
cand_errors_teamwise <- cand_errors_teamwise[order(cand_errors_teamwise$tl_name, cand_errors_teamwise$date, decreasing = F),]

cand_errors_teamwise_agg <- as.data.frame(cand_errors_agg %>% group_by(tl_name) %>%
                                          summarise(no_of_days = length(unique(date)),
                                                    district = sum(district == 0),
                                                    ac = sum(ac == 0),
                                                    gender = sum(gender == 0),
                                                    age = sum(age == 0),
                                                    party_vote_2019_election = sum(party_vote_2019_ls_elections == 0),
                                                    party_vote_upcoming = sum(party_vote_upcoming_bypoll == 0)))

cand_errors_teamwise_agg$agg_total_errors <- rowSums(cand_errors_teamwise_agg[3:length(cand_errors_teamwise_agg)])
cand_errors_teamwise_agg$errors_per_day <- round(cand_errors_teamwise_agg$agg_total_errors/cand_errors_teamwise_agg$no_of_days,2) 
# gen_errors_teamwise_agg <- unique(merge(gen_errors_agg[,c("agent_id","agent_name","tl_name")], gen_errors_agentwise_agg, by = "agent_id", all.y = T))
# gen_errors_teamwise_agg <- as.data.frame(gen_errors_agentwise_agg %>% group_by(agent_id) %>% mutate(rank = 1:length(agent_name)) %>% filter(rank == 1))
# gen_errors_agentwise_agg$rank <- NULL
cand_errors_teamwise_agg <- cand_errors_teamwise_agg[order(cand_errors_teamwise_agg$errors_per_day, decreasing = T),]

cand_error_wb <- createWorkbook("cand_error_tracker")
daily_cand_errors_agentwise <- read_excel("cand_error_tracker.xlsx", sheet = 1)
daily_cand_errors_agentwise$date <- as.character(as.POSIXct(daily_cand_errors_agentwise$date, format = "%Y-%m-%d"))
cand_errors_agentwise$date <- as.character(cand_errors_agentwise$date)
daily_cand_errors_agentwise <- rbind(daily_cand_errors_agentwise,cand_errors_agentwise)
# daily_cand_errors_agentwise <- cand_errors_agentwise
addWorksheet(cand_error_wb, sheetName = "agent_level")
writeData(cand_error_wb, "agent_level", daily_cand_errors_agentwise, rowNames = F)

daily_cand_errors_teamwise <- read_excel("cand_error_tracker.xlsx", sheet = 2)
daily_cand_errors_teamwise$date <- as.character(as.POSIXct(daily_cand_errors_teamwise$date, format = "%Y-%m-%d"))
cand_errors_teamwise$date <- as.character(cand_errors_teamwise$date)
daily_cand_errors_teamwise <- rbind(daily_cand_errors_teamwise,cand_errors_teamwise)
# daily_cand_errors_teamwise <- cand_errors_teamwise
addWorksheet(cand_error_wb, sheetName = "team_level")
writeData(cand_error_wb, "team_level", daily_cand_errors_teamwise, rowNames = F)

addWorksheet(cand_error_wb, sheetName = "agent_level_aggregate")
writeData(cand_error_wb, "agent_level_aggregate", cand_errors_agentwise_agg, rowNames = F)

addWorksheet(cand_error_wb, sheetName = "team_level_aggregate")
writeData(cand_error_wb, "team_level_aggregate", cand_errors_teamwise_agg, rowNames = F)


saveWorkbook(cand_error_wb, file = "cand_error_tracker.xlsx", overwrite = T)

