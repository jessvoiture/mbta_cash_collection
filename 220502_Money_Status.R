# Cash Collection Schedule
# Last Modified: 5/2/22
# Reads in daily reports and generates daily schedule, burn rates, growth rates, and all data.

# packages ----------------------------------------------------------------

library(tidyverse, quietly = T)
library(dplyr, quietly = T)
library(readr, quietly = T)
library(purrr, quietly = T)
library(broom)
library(lubridate, quietly = T)
library(readxl, quietly = T)
library(openxlsx, quietly = T)
library(magrittr, quietly = T)
library(data.table, quietly = T)

# file imports ------------------------------------------------------------
## set working directory where ALL files are located (use no sub-directories)
setwd("~/Documents/School/mbta/Cash Collection/Money Status Burn Rates/")

# path to all consolidated sheets and daily reports
path <- "~/Documents/School/mbta/Cash Collection/Money Status Burn Rates/"

# import and merge consolidated sheets
## set path where ALL files are located (use no sub-directories)
consolidated_path <- paste(path, "Consolidated_report/consolidated_report.csv", sep = "")
df <- read_csv(consolidated_path)

# import cash on hand (coh) and stacker DAILY reports
daily_report_path <- paste(path, "daily_reports/", sep = "")

coh_pattern <- paste("cash-on-hand.", as.character(format(Sys.Date(), "%y%m%d")), "*.csv", sep="")
stacker_pattern <- paste("stacker.", as.character(format(Sys.Date(), "%y%m%d")), "*.csv", sep="")

coh_files <- list.files(daily_report_path, pattern=glob2rx(coh_pattern), full.names=T, recursive=F) 
df_coh <- lapply(coh_files, function(x) {
  csv <- read.csv(x, header = T, sep=",")[-c(1), ] # need to skip the first row which also contains headers
  csv <- csv %>%
    mutate(file_name = basename(x))
}) %>% 
  bind_rows()

stacker_files <- list.files(daily_report_path, pattern=glob2rx(stacker_pattern), full.names=T, recursive=F) 
df_stack <- lapply(stacker_files, read.csv) %>%
  bind_rows()

# import service issues reports: receipt roll errors (farego) and workorders (servicenow)
# receipt rolls
rr_path <- paste(path, "receipts/", sep="")
rr_file_info <- file.info(list.files(rr_path, pattern=glob2rx("event-detail.*.csv"), full.names = T))
newest_rr <- rownames(rr_file_info)[which.max(rr_file_info$mtime)] # selects most recent file

# need to skip the first row which also contains headers
df_rr_raw <- read.csv(newest_rr, header = T, sep=",")[-c(1), ]

# workorders
wo_path <- paste(path, "work_orders/", sep="")

wo_file_info <- file.info(list.files(wo_path, pattern=glob2rx("wm_order*.xlsx"), full.names = T))
newest_wo <- rownames(wo_file_info)[which.max(wo_file_info$mtime)] # selects most recent file

df_wo_raw <- read_xlsx(newest_wo)

# file to rename station names (servicenow has epithetical station names)
rename_wo_station <- read_xlsx("work_orders/work_orders_rename.xlsx")

# service days with corresponding dates (until end of 2023)
service_day_dates_path <- paste(path, "service_days_info/Service_day_dates.csv", sep="")
service_day_dates <- read.csv(service_day_dates_path)

# file with Devices and corresponding service day
devs_service_day_path <- paste(path, "service_days_info/Device_service_days.csv", sep="")
df_info <- read.csv(devs_service_day_path)

# constant definitions ----------------------------------------------------

# subset for useful fields from daily coh reports
good_coh_cols <- c("Station", "Device", "Date", "nickle", "qtr", "dollar", 
                   "cv", "bnv", "notes")

# subset for useful fields from daily stacker reports
good_stack_cols <- c("Station", "Device", "Date", "stack5", "stack6",
                     "lum9", "lum10", "lum13", "lum14")

# subset for useful fields from daily receipt roll error report
good_rr_cols <- c("Locality", "X", "Date", "Event.Description")

# subset for useful fields from work order report
good_wo_cols <- c("Station", "Date", "Device", "container", "Short_Descr", "Issue", "Desc", "Date_service")

# subset for coh columns that need to be converted to num
coh_cols_to_convert <- c("nickle", "qtr", "dollar", "cv", "bnv", "one_dollar",
                         "two_dollar", "five_dollar", "ten_dollar", "twenty_dollar",
                         "fifty_dollar", "hundred_dollar")

# subset for container columns
container_cols <- c("nickle", "qtr", "dollar", "cv", "bnv", "notes", 
                    "stack5", "stack6", "lum9", "lum10", "lum13", "lum14")

# container list
consumables_list <- c("stack5", "stack6", "lum9", "lum10", "lum13", "lum14")
stacks_list <- c("stack5", "stack6")
lums_list <- c("lum9", "lum10", "lum13", "lum14")
coins_list <- c("nickle", "qtr", "dollar")

# container column order in the excel sheet
desired_cointainer_order <- c("nickle", "qtr", "dollar", "cv", "bnv", "notes",
                              "stack5", "stack6", "lum9", "lum10", "lum13", "lum14")

# subset for info fields
info_cols <- c("Service_Day", "Station", "Device")

# Service Day identifiers  
every_other_week_names <- c("Mon. A", "Mon. B", "Tue. A", "Tue. B", "Wed. A", 
                            "Wed. B", "Thur. A", "Thur. B", "Fri. A", "Fri. B")

every_week_names <- c("Mon", "Tue", "Wed", "Thur", "Fri")

# old and new column name lists
coh_rename_cols <- c("Station" = "Station.Name", "nickle" = "Coins.in.Hopper","qtr" = "X.1", 
                     "dollar" = "X.3", "cv" = "CV.Amount....", "bnv" = "BNV.Amount....",
                     "one_dollar" = "Bills.in.BNV", "two_dollar" = "X.12", 
                     "five_dollar" = "X.13", "ten_dollar" = "X.14", "twenty_dollar" = "X.15", 
                     "fifty_dollar" = "X.16", "hundred_dollar" = "X.17")

stack_rename_cols <- c("Station" = "Locality", "stack5" = "RemStacker5", 
                       "stack6" = "RemStacker6", "lum9" = "RemStacker9", "lum10"  = "RemStacker10", 
                       "lum13" = "RemStacker13", "lum14" = "RemStacker14")

rr_rename_cols <- c("Station" = "Locality", "Device" = "X", "Desc" = "Event.Description")

# thresholds
hopper_capacity <- 1000
cv_capacity <- 350
bnv_capacity <- 12000
notes_capacity <- 1200
stack_capacity <- 325
lum_capacity <- 1400
pool_stack_capacity <- 750
pool_lum_capacity <- 5600

# filter dates
three_wks_ago <- 21

# dates
todays_date <- Sys.Date()
todays_month <- month.name[month(todays_date)]

# export variables --------------------------------------------------------

# csv
# daily report
subDir <- paste(todays_date, "_money_status", sep="")

# check if directory already exists; if not then create it
ifelse(!dir.exists(file.path(path, subDir)), 
       dir.create(file.path(path, subDir)), 
       "Directory already exists")

# consolidated monthly report
subDir_month <- paste("monthly_money_status/", todays_month, "_money_status", sep="")

# check if directory already exists; if not then create it
ifelse(!dir.exists(file.path(path, subDir_month)), 
       dir.create(file.path(path, subDir_month)), 
       "Directory already exists")

# consolidated report (Jan 1 2022 - present)
subDir_all <-"consolidated_report/"

# check if directory already exists; if not then create it
ifelse(!dir.exists(file.path(path, subDir_all)), 
       dir.create(file.path(path, subDir_all)), 
       "Directory already exists")

export_path <- paste(path, subDir, "/", sep="")
export_path_monthly <- paste(path, subDir_month, "/", sep="")
export_path_all_data <- paste(path, subDir_all, "/", sep="")

# output names
output_name_raw <- paste(export_path, todays_date, "_money_status", ".xlsx", sep="") # today's file
output_name_month <- paste(export_path_monthly, todays_month, "_money_status", ".xlsx", sep="") # monthly file
output_name_all_data <- paste(export_path_all_data, "consolidated_report", ".csv", sep="") # consolidated file
output_name_tableau <- paste(path, "money_status1", ".csv", sep="") # tableau file
output_name_stats <- paste(export_path, todays_date, "_stats", ".csv", sep="")

# functions ---------------------------------------------------------------

# purpose of function is to remove extreme negative outliers
remove_outliers <- function(x, na.rm = TRUE, ...) {
  sdx <- sd(x, na.rm = T)
  if (sdx <= 3 | is.na(sdx)) { # higher daily standard deviation indicates extreme values (3% is arbitrary; however sds above 3 appear to have extreme values in the dataset)
    x
  } else  {
    meanx <- mean(x, na.rm = T)
    sdfactor <- 2 # 2 standard deviations below mean (95%); can increase with more data points
    y <- x
    y[x < (meanx - sdfactor * sdx)] <- NA
    y
  }
}

find_matched_rows <- function(df1, df2) {
  df1_transposed <- data.table::transpose(df1)
  df2_as_vector <- unlist(df2)
  match_map <- lapply(df1_transposed,FUN = `%in%`, df2_as_vector) %>%
    as.data.frame(stringsAsFactors = FALSE) %>%
    sapply(function(x) sum(x) > 0)
  matched_rows <- seq(1:nrow(df1))[match_map]
  return(list(matched_rows))
}

# data wrangling ----------------------------------------------------------

# coh data
df_coh <- df_coh %>%
  rename(all_of(coh_rename_cols)) %>%
  filter(Station != "Station Name")  %>%
  filter(nickle != "5c")  %>%
  mutate(across(all_of(coh_cols_to_convert), parse_number)) %>%
  # raw data is broken into bill denominations; line below sums them into one notes count
  mutate(notes = one_dollar + two_dollar + twenty_dollar + fifty_dollar + hundred_dollar) %>%
  # data does not contain date field; date is found in 6 characters after "cash-on-hand.*" in filename
  mutate(Date_str = sub("cash-on-hand.(\\d{6}).*", "\\1", file_name)) %>%
  mutate(Date = ymd(Date_str)) %>%
  select(all_of(good_coh_cols)) %>%
  filter(!(is.na(nickle))) %>%
  mutate(
    Station = recode(
      Station,
      'Chestnut Hills' = 'Chestnut Hill',
      'Community ' = "Community College",
      'Dudley Station' = 'Nubian',
      'Copley Square' = 'Copley',
      'Hynes' = 'Hynes Convention Center',
      'Kenmore Square' = 'Kenmore',
      'Government' = 'Government Center',
      'Downtown' = 'Downtown Crossing',
      'Tufts Medical' = 'Tufts Medical Center'
    )
  )

# stacker data
df_stack <- df_stack %>%
  rename(all_of(stack_rename_cols)) %>%
  mutate(Device = as.character(DevID)) %>%
  # synchronization is the date-time field on stacker report
  mutate(Date = as.Date(Synchronization, format = "%m/%d/%Y")) %>%
  select(all_of(good_stack_cols))  %>%
  mutate(
    Station = recode(
      Station,
      'Chestnut Hills' = 'Chestnut Hill',
      'Community ' = "Community College",
      'Dudley Station' = 'Nubian',
      'Copley Square' = 'Copley',
      'Hynes' = 'Hynes Convention Center',
      'Kenmore Square' = 'Kenmore',
      'Government' = 'Government Center',
      'Downtown' = 'Downtown Crossing',
      'Tufts Medical' = 'Tufts Medical Center'
    )
  )

# merge stacker and coh report
df_daily <- right_join(df_coh, df_stack, by = c("Device" = "Device",
                                                "Date" = "Date",
                                                "Station" = "Station"))

# dataframe for information columns
df_info <- df_info %>%
  mutate(Device = as.character(Device))

# service day dates
service_day_dates <- service_day_dates %>%
  pivot_longer(c(Weekly, Every_other, biweekly, triweekly),
               names_to = "Service_Day",
               values_to = "Day") %>%
  select(-Service_Day) %>%
  rename("Future_date" = "Date",
         "Service_Day" = "Day") %>%
  mutate(Future_date = as.Date(Future_date, format = "%m/%d/%y")) %>%
  filter(Future_date >
           ifelse(hour(Sys.time()) > 9, todays_date+ 1, todays_date))  %>%
  group_by(Service_Day) %>%
  summarise(Next_Service_Date = min(Future_date))

# receipt rolls
df_rr <- df_rr_raw %>%
  select(all_of(good_rr_cols)) %>%
  rename(all_of(rr_rename_cols)) %>%
  distinct(Device, .keep_all = T) %>%
  mutate(container = "RR") %>%
  mutate(Date = as.Date(Date, format =  "%m/%d/%Y")) %>%
  mutate(Issue = "RR") %>%
  mutate(Date_service = todays_date)

rr_for_daily_report <- df_rr %>%
  select(Device, Date, container) %>%
  rename("RR" = "container")

# work orders
# recode column names and add standard station names to replace old ones
df_wo <- df_wo_raw %>%
  rename("Symptom_code" = "Symptom Code", "Device" = "Device ID", 
         "Short_Desc" = "Short description", "Desc" = "Detailed description") %>%
  left_join(rename_wo_station, by="Location") %>%
  select(Station, Device, Short_Desc, Desc)

# slice up short desc column and reformat date
df_wo <- df_wo %>%
  mutate(Short_Descr = Short_Desc) %>% # make a duplicate of column and split one of them up into components
  # Short_Desc contains assignment group, date of service, and the issue (ex: Brinks 3/21 ACM1)
  # if there are additional items after the first issue, they are added to the issue column
  separate(
    Short_Desc,
    into = c('Brinks', 'Date_str', "Issue"),
    sep = " ",
    extra = "merge"
  ) %>%
  separate(Date_str, into = c("Month", "Day"), sep = "/") %>%
  mutate(Year = "2022") %>%
  mutate(Date_service = make_datetime(Year, Month, Day)) %>%
  select(-c(Month, Day, Year))

# add columns 
df_wo <- df_wo %>%
  mutate(Date = todays_date) %>%
  mutate(container = "wo") %>%
  group_by(Device) %>%
  summarise(across(everything(), str_c, collapse="")) %>%
  select(Station, Device, Date, Short_Descr, Desc, container, Date_service) %>%
  rename("Issue" = "Short_Descr")

wo_for_daily_report <- df_wo %>%
  select(Device, Issue) %>%
  rename("wo" = "Issue")

# union rr and wo
df_service_issues <- union(df_wo, df_rr)

# merge daily reports and consolidated reports
df <- df %>%
  select(-Service_Day) %>%
  mutate(Device = as.character(Device))

df <- bind_rows(df, df_daily) %>%
  distinct(Station, Date, Device, .keep_all = T)

max_date <- max(df$Date, na.rm = T)
max_month <- max(month(df$Date), na.rm = T)

# "^221" encodes for cash devices
# na for FVM indicates no notes; NA for CLFVM indicates not applicable (machines 
# do not accept bills)
# replace NA in FVM notes column (Device starts with 221) with 0
df <- df %>%
  mutate(notes = replace(notes, is.na(notes) & str_detect(Device, "^221"), 0))

df <- df %>%
  select(Station, Date, Device, all_of(container_cols)) 

df_this_month_export <- df %>% # just this month's data
  filter(month(Date) == month(todays_date)) %>%
  left_join(df_info, by = c("Station", "Device")) %>%
  relocate(Service_Day)

df_today_export <- df_this_month_export %>% # just today's data
  filter(Date == todays_date) %>%
  left_join(rr_for_daily_report, by = c("Device", "Date")) %>%
  left_join(wo_for_daily_report, by = c("Device")) %>%
  arrange(Station, Device)

df_today_fvm <- df_today_export %>% # today's FVM data
  filter(str_detect(Device, "^221"))

df_today_clfvm <- df_today_export %>%
  filter(str_detect(Device, "^222")) %>% # today's CLFVM data
  select(Service_Day, Station, Date, Device, all_of(consumables_list), RR, wo) # subset to relevant cols

df_all <- df %>%
  left_join(df_info, by = c("Station", "Device")) %>%
  relocate(Service_Day)

# export data before any further manipulation as it reflects current money status file
write.xlsx(df_this_month_export, output_name_month, overwrite = T) # just this month's data

todays_data_list <- list("FVM" = df_today_fvm, "CLFVM (cashless)" = df_today_clfvm) # today's data with a tab each for FVM and CLFVM
write.xlsx(todays_data_list, file = output_name_raw, overwrite = T, append = T)

write.csv(df_all, output_name_all_data, row.names = F)

# reorient bank vault (bnv), notes, and coin vault (cv)
# currently bnv, notes, and cv go up as people insert money into machines
# reorientation will reflect the remaining capacity
df <- df %>%
  pivot_longer(all_of(container_cols),
               names_to = "container",
               values_to = "level_raw") %>%
  mutate(
    level = case_when(
      container == "bnv" ~ bnv_capacity - level_raw,
      container == "cv" ~ cv_capacity - level_raw,
      container == "notes" ~ notes_capacity - level_raw,
      TRUE ~ level_raw
    )
  )

# remove instances where measured capacity is above allowed capacity
# this occurs when the machine is set to the incorrect amount and results in
# large daily differences 
df <- df %>%
  filter(
    level <= case_when(
      container %in% coins_list ~ hopper_capacity,
      container == "bnv" ~ bnv_capacity,
      container == "cv" ~ cv_capacity,
      container == "notes" ~ notes_capacity,
      container %in% stacks_list ~ stack_capacity,
      container %in% lums_list ~ lum_capacity
    )
  )

# calculate daily difference
df <- df %>%
  group_by(Device, container) %>%
  arrange(Device, container, Date, .by_group = T) %>%
  mutate(level_diff = level - lag(level))

# remove instances where container has reached empty over several days
df <- df %>%
  mutate(
    level_diff = case_when(
      container %in% consumables_list ~
        # lums and stacks stop dispensing at 10; therefore empty = 10
        ifelse(level_diff == 0 & level == 10, NA_real_, level_diff),
      container %in% coins_list ~
        # hoppers empty until they reach 0
        ifelse(level_diff == 0 & level == 0, NA_real_, level_diff),
      TRUE ~ level_diff
    )
  )

# calculate percents
df <- df %>%
  mutate(
    pct_remaining = 100 * level /
      case_when(
        container %in% coins_list ~ hopper_capacity,
        container == "bnv" ~ bnv_capacity,
        container == "cv" ~ cv_capacity,
        container == "notes" ~ notes_capacity,
        container %in% stacks_list ~ stack_capacity,
        container %in% lums_list ~ lum_capacity,
      )
  ) %>%
  # calculate daily pct difference
  group_by(Device, container) %>%
  arrange(Device, container, Date, .by_group = T) %>%
  mutate(pct_diff = pct_remaining - lag(pct_remaining))

# remove outliers
df <- df %>%
  group_by(Device, container) %>%
  mutate(level_diff_clean = remove_outliers(level_diff)) %>%
  mutate(pct_diff_clean = remove_outliers(pct_diff)) %>%
  arrange(Device, container, Date, .by_group = T) %>%
  select(-c(pct_diff, level_diff)) 

# burn rate calculations ------------------------------------------------------------

df_stat <- df %>%
  # burn rates change seasonally, due to current events, etc; therefore use most
  # recent 3 weeks to calculate burn rates
  filter(Date >= max_date - three_wks_ago) %>%
  # positive values reflect machine refills
  filter(pct_diff_clean <= 0) %>%
  group_by(Device, container) %>%
  summarise(
    daily_mean_pct = abs(mean(pct_diff_clean, na.rm = T)),
    daily_sd_pct = sd(pct_diff_clean, na.rm = T),
    daily_mean_level = abs(mean(level_diff_clean, na.rm = T)),
    daily_sd_level = sd(level_diff_clean, na.rm = T)
  )

# join data frames (main df, service issues, and info) ------------------

df <- df %>%
  left_join(df_stat, by = c("Device" = "Device",
                            "container" = "container")) %>%
  bind_rows(df_service_issues) %>%
  left_join(df_info, by = c("Station" = "Station",
                            "Device" = "Device"))

write.csv(df, output_name_tableau)

# growth rate calculation -------------------------------------------------

df_growth_rate <- df %>%
  mutate(week_num = week(Date)) %>%
  filter(week_num < max(week_num)) %>%
  filter(level_diff_clean <= 0) %>%
  group_by(Device, container, week_num) %>%
  summarise(weekly_diff = sum(level_diff_clean)) %>% # amount used in 1 week
  ungroup() %>%
  mutate( # pct used in 1 weeke
    weekly_pct_diff = -100 * weekly_diff / 
      case_when(
        container %in% coins_list ~ hopper_capacity,
        container == "bnv" ~ bnv_capacity,
        container == "cv" ~ cv_capacity,
        container == "notes" ~ notes_capacity,
        container %in% stacks_list ~ stack_capacity,
        container %in% lums_list ~ lum_capacity,
        container == "pooled_stack" ~ pool_stack_capacity,
        container == "pooled_lum" ~ pool_lum_capacity
      )
  )

# calculate week over week growth (pct)
df_growth_rate <- df_growth_rate %>%
  group_by(Device) %>%
  mutate(wow_growth_pct = (weekly_pct_diff - lag(weekly_pct_diff)) / lag(weekly_pct_diff) * 100)

# fit linear model
df_model <- df_growth_rate %>%
  group_by(Device, container) %>%
  filter(week_num > max(week_num) - 6) %>% # filter to last 6 weeks for regression
  do(tidy(lm(weekly_pct_diff ~ week_num, .))) %>%
  filter(term == "week_num") %>%
  select(Device, container, estimate) %>%
  rename("growth_6wks" = "estimate")

df_growth_rate <- df_growth_rate %>%
  filter(week_num > max(week_num) - 4) %>% # only show week over week growth for past 4 weeks
  left_join(df_info, by = "Device") %>%
  left_join(df_model, by = c("Device", "container")) %>%
  select(Station, Device, Service_Day, container, week_num, weekly_diff, weekly_pct_diff, wow_growth_pct, growth_6wks)

write.csv(df_growth_rate, output_name_stats, row.names = F)

# threshold calculations --------------------------------------------------

# based on number of days until the next service and burn rate, filter to devices
# that require service
df_need_service <- df %>%
  filter(Date == max_date) %>%
  left_join(service_day_dates, by = "Service_Day") %>%
  mutate(days_to_service = as.numeric(Next_Service_Date - Date)) %>%
  mutate(need_service = ifelse(level < 
                                 case_when(
                                   container %in% coins_list ~ days_to_service * daily_mean_level,
                                   container == "cv" ~ days_to_service * daily_mean_level + 50, # 50 = cv buffer
                                   # additional buffer for bnv/notes to desired threshold
                                   container == "bnv" ~ days_to_service * daily_mean_level + 5000, # 5000 = bnv buffer
                                   container == "notes" ~ days_to_service * daily_mean_level + 500, # 500 = notes buffer
                                   # less buffer for consumables since there are multiple of them
                                   container %in% consumables_list ~ 40), "yes", "no")) %>% # 40 = consumables threshold
  select(Date, Station, Device, Service_Day, container, level_raw, need_service) %>%
  filter(need_service == "yes") %>%
  select(Device, container) %>%
  ungroup()

fvm_need_service <- df_need_service %>%
  filter(str_detect(Device, "^221")) # 221 at the beginning of the Dev ID indicates a FVM

clfvm_need_service <- df_need_service %>%
  filter(str_detect(Device, "^222")) # 222 at the beginning of the Dev ID indicates a CLFVM

# split data frame by container type
# output is a list of vectors, one for each container type
# purpose is so that each container list corresponds to a column in the excel workbook; therefore
# we can iterate through each container type and each column
fvm_need_service <- fvm_need_service %>%  
  group_by(container) %>%
  group_nest() %>%
  arrange(match(container, desired_cointainer_order)) # sort containers as they appear in the excel file

clfvm_need_service <- clfvm_need_service %>%  
  group_by(container) %>%
  group_nest() %>%
  arrange(match(container, desired_cointainer_order))

# initialize array
fvm_matched_rows <- data.frame()
clfvm_matched_rows <- data.frame()

# find the values in today's money status that have containers that need service
# each tibble in df_need_service corresponds to a container
for (i in 1:nrow(fvm_need_service)) {
  fvm_matched_rows[i, 1] <- fvm_need_service$container[i] # container type
  fvm_matched_rows[i, 2] <- # index (row number) of machines that need service
    # df_today_export = today's data for all machines
    # fvm_need_service = which containers and which machines require service for fvm machines
    list(find_matched_rows(df_today_fvm, fvm_need_service$data[i]))
}

for (i in 1:nrow(clfvm_need_service)) {
  clfvm_matched_rows[i, 1] <- clfvm_need_service$container[i] # container type
  clfvm_matched_rows[i, 2] <- # index (row number) of machines that need service
    # df_today_export = today's data for all machines
    # fvm_need_service = which containers and which machines require service for fvm machines
    list(find_matched_rows(df_today_clfvm, clfvm_need_service$data[i]))
}

# assign containers numbers that correspond to their column order in the excel file
# addStyle function accepts integers for cols/rows field, so necessary to translate 
# container string to integer
fvm_container_col_num <- data.frame(container = container_cols,
                                    col_num = 1:length(container_cols))

clfvm_container_col_num <- data.frame(container = consumables_list,
                                      col_num = 1:length(consumables_list))

fvm_matched_rows <- fvm_matched_rows %>%
  rename("container" = "V1",
         "matched_rows" = "V2") %>%
  left_join(fvm_container_col_num, by = "container")

clfvm_matched_rows <- clfvm_matched_rows %>%
  rename("container" = "V1",
         "matched_rows" = "V2") %>%
  left_join(clfvm_container_col_num, by = "container")

## create workbook with condition highlighting
highlight_style <- createStyle(fgFill = "yellow") 

today_money_status_wb <-loadWorkbook(file = output_name_raw)

for (i in 1:nrow(fvm_matched_rows)) {
  addStyle(
    wb = today_money_status_wb,
    sheet = 1, # FVM sheet is first
    style = highlight_style,
    rows = 1 + unlist(fvm_matched_rows$matched_rows[i]), # 1 is header row
    cols = 4 + fvm_matched_rows$col_num[i], # 4 is the information columns (Service Day, Date, Station, Device)
    stack = TRUE,
    gridExpand = TRUE
  )
}

for (i in 1:nrow(clfvm_matched_rows)) {
  addStyle(
    wb = today_money_status_wb,
    sheet = 2, # CLFVM sheet is second
    style = highlight_style,
    rows = 1 + unlist(clfvm_matched_rows$matched_rows[i]), # 1 is header row
    cols = 4 + clfvm_matched_rows$col_num[i], # 4 is the information columns (Service Day, Date, Station, Device)
    stack = TRUE,
    gridExpand = TRUE
  )
}

saveWorkbook(wb = today_money_status_wb,
             file = output_name_raw,
             overwrite = TRUE)

