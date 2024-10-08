# Path to the Excel file
file_path <- paste0(input_file_path, "/BSMHFT Sitrep")

# Access the Excel file within the BSMHFT Sitrep New data folder
excel_file <- list.files(path = paste0(file_path, "/New data"), full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_BSMHFT_Sitrep_Test.xlsx")

# Read raw Excel file
data <- read_xlsx(excel_file[1])

# Remove unnecessary rows & columns
data_2 <- data %>% 
  janitor::row_to_names(row_number = which(data$`BSMHFT Daily Mental Health SitRep` == "No.")) %>%
  select(-c(`No.`, Definition))

# Initialize an empty list to store data frames 
data_list <- list()


#1. Convert Excel & POSIXct date columns to human-readable dates---------------

convert_dates_cols <- function(data) {
  
  #1. POSIXct column names conversion
  posixct_cols <- sapply(data, function(x) inherits(x, "POSIXct"))
  posixct_col_names <- names(data)[posixct_cols]
  
  # New POSIXct column names initialization
  new_posixct_col_names <- vector("character", length(posixct_col_names))
  
  for (i in seq_along(posixct_col_names)) {
    posix_date <- as.POSIXct(as.numeric(posixct_col_names[i]), origin = "1970-01-01", tz = "UTC")
    new_posixct_col_names[i] <- format(posix_date, format = "%d/%m/%Y")
  }
  
  # Update column names for POSIXct
  names(data)[posixct_cols] <- new_posixct_col_names
  
  #2. Excel dates columns conversion
  # Directly identifying columns assumed to be Excel dates, excluding specific names
  excel_date_candidate_cols <- !(names(data) %in% new_posixct_col_names) & 
    !(names(data) %in% c("Metric", "BSMHFT Lead"))
  excel_dates_col_names <- names(data)[excel_date_candidate_cols]
  
  # New Excel date column names initialization
  new_excel_dates_col_names <- vector("character", length(excel_dates_col_names))
  
  for (i in seq_along(excel_dates_col_names)) {
    
    excel_date <- as.Date(as.numeric(excel_dates_col_names[i]), origin = "1899-12-30")
    new_excel_dates_col_names[i] <- format(excel_date, format = "%d/%m/%Y")
  }
  
  # Update column names for Excel dates
  names(data)[excel_date_candidate_cols] <- new_excel_dates_col_names
  
  return(data)
}

#2. Process simple metrics ----------------------------------------------------

#1. MFFD Transferred Out
#2. MH Support for older people stepped down into P2
#3. 12 hour breaches in ED
#4. Total awaiting admission in community
#5. PDU referrals
#6. POS referrals
#7. POS Transfers from ED

process_data_v1 <- function(data, metric_name, BSMHFT_lead){
  
  data %>% 
    filter(Metric == metric_name) %>% 
    convert_dates_cols() %>% 
    mutate(`BSMHFT Lead` = BSMHFT_lead) %>% 
    mutate(
      across(.cols = 3:ncol(data_2),
             .fns = ~ ifelse(is.na(.) | !str_detect(., "^[0-9.]+$"), NA_real_, as.numeric(.)))) %>% 
    pivot_longer(
      cols = -c('Metric', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value") %>% 
    mutate(
      Value = ifelse(is.na(Value), NA_real_, as.numeric(Value)),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP") %>% 
    mutate(`Metric Category Type` = "Total",
           `Metric Category Name` = "Total",
           Metric = metric_name,
           Provider = NA) %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
}

# A data frame of metrics and corresponding BSMHFT leads to process
metrics_leads <- data.frame(
  Metric = c(
    "MFFD Transferred Out",
    "MH Support for older people stepped down into P2",
    "12 hour breaches in ED",
    "Total awaiting admission in community",
    "PDU referrals",
    "POS referrals",
    "POS Transfers from ED"
  ),
  BSMHFT_Lead = c(
    "Bed Management",
    NA,
    "Informatics",
    "Bed Management",
    "Bed Management",
    "Bed Management",
    "Bed Management"
  )
)

# Use lapply to iterate over each row of metrics_leads and update data_list
# This checks if the metric name already exists in data_list; if not, it adds the new processed data
for (i in 1:nrow(metrics_leads)) {
  metric_name <- metrics_leads$Metric[i]
  BSMHFT_lead <- metrics_leads$BSMHFT_Lead[i]
  
  # Only add the new processed data if the metric name does not already exist in data_list
  if (!metric_name %in% names(data_list)) {
    data_list[[metric_name]] <- process_data_v1(data = data_2,
                                                metric_name = metric_name,
                                                BSMHFT_lead = BSMHFT_lead)
  }
}

# data_list[["MH Support for older people stepped down into P2"]] %>%
#   filter(Date == "19/04/2022") %>%
#   View()

#2. Process metrics with providers only ----------------------------------------

#1. Total ED awaiting assessment
#2. Length of wait for ED assessment
#3. % seen within 1 hour assessment target in ED
#4. S136 patients in ED

process_data_v2 <- function(data, row_value_start, row_value_end, 
                            BSMHFT_Lead, metric_name){
  start_row <- which(data$Metric == row_value_start)[1]
  end_row <- which(data$Metric == row_value_end)[1]
  
  data %>% 
    slice(start_row: (end_row - 1)) %>% 
    convert_dates_cols() %>% 
    mutate(`BSMHFT Lead` = BSMHFT_Lead) %>%
    as_tibble() %>% 
    mutate(
      across(.cols = 3:ncol(data),
             .fns = ~ifelse(is.na(.), NA_real_, as.numeric(.)))) %>%
    pivot_longer(
      cols = -c('Metric', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value") %>% 
    mutate(
      Value = as.numeric(Value),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP") %>%
    mutate(Provider = case_when(
      grepl("City", Metric) ~ "City",
      grepl("Good Hope", Metric) ~ "Good Hope",
      grepl("Heartlands", Metric) ~ "Heartlands",
      grepl("QE/UHB", Metric) ~ "QE/UHB",
      TRUE ~ "System" )) %>%
    mutate(`Metric Category Type` = "Total",
           `Metric Category Name` = "Total",
           Metric = metric_name) %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
}


# Save the data frames in a data list
data_list[["Total ED awaiting assessment"]] <- process_data_v2(data = data_2,
                                                               row_value_start = "Total ED awaiting assessment:",
                                                               row_value_end = "Length of wait for ED assessment:",
                                                               BSMHFT_Lead = "Bed Management", 
                                                               metric_name = "Total ED awaiting assessment")

data_list[["Length of wait for ED assessment"]] <- process_data_v2(data = data_2,
                                                                   row_value_start = "Length of wait for ED assessment:",
                                                                   row_value_end = "% seen within 1 hour assessment target in ED:",
                                                                   BSMHFT_Lead = "Informatics", 
                                                                   metric_name = "Length of wait for ED assessment")  

data_list[["S136 patients in ED"]] <- process_data_v2(data = data_2,
                                                      row_value_start = "S136 patients in ED:",
                                                      row_value_end = "12 hour breaches in ED",
                                                      BSMHFT_Lead = "Bed Management", 
                                                      metric_name = "S136 patients in ED") 


# % seen within 1 hour assessment target 

start_row <- which(data_2$Metric == "% seen within 1 hour assessment target in ED:")[1]
end_row <- which(data_2$Metric == "Total IP MH:")[1] - 1

pct_seen_within_target <- data_2 %>% 
  slice(start_row : end_row) %>% 
  convert_dates_cols()

## Process columns with mixed numbers and percentages separately
pct_cols_to_modify <- pct_seen_within_target %>% 
  select(where(is.character)) %>% 
  pivot_longer(
    cols = -c(Metric, `BSMHFT Lead`),
    names_to = "Date",
    values_to = "Value"
  ) %>% 
  # Extract numbers within brackets and % by matching string characters
  # Store these as denominators and %, from which numerators are derived
  mutate(
    Denominator = ifelse(is.na(Value), NA, as.numeric(str_match(Value, "\\((\\d+)\\)")[,2])),
    Percentage = ifelse(is.na(Value), NA, as.numeric(str_match(Value, "(\\d+)%")[,2])),
    Percentage = Percentage/100,
    Numerator = as.integer(Percentage * Denominator)
  ) %>% 
  select(Metric, `BSMHFT Lead`, Date, Numerator, Denominator, Percentage) %>% 
  pivot_longer(
    cols = -c(Metric, `BSMHFT Lead`, Date),
    names_to = "Metric Category Type",
    values_to = "Value"
  ) %>% 
  select(Metric, `Metric Category Type`, `BSMHFT Lead`, Date, Value)

## Combine with rest of data and process as normal
pct_seen_within_target <- pct_seen_within_target %>% 
  as_tibble() %>% 
  select(Metric, `BSMHFT Lead`, where(is.numeric)) %>% 
  pivot_longer(
    cols = -c('Metric', 'BSMHFT Lead'),
    names_to = "Date",
    values_to = "Value") %>% 
  mutate(`Metric Category Type` = "Percentage") %>% 
  select(Metric, `Metric Category Type`, `BSMHFT Lead`, Date, Value) %>% 
  bind_rows(pct_cols_to_modify) %>% 
  mutate(
    Value = as.numeric(Value),
    Day = wday(Date, label = TRUE, abbr = FALSE),
    `File Name` = "BSMHFT Daily MH SITREP") %>% 
  mutate(Provider = case_when(
    grepl("City", Metric) ~ "City",
    grepl("Good Hope", Metric) ~ "Good Hope",
    grepl("Heartlands", Metric) ~ "Heartlands",
    grepl("QE/UHB", Metric) ~ "QE/UHB",
    TRUE ~ "System" )) %>%
  mutate(
    Metric = "% seen within 1 hour assessment target in ED",
    `BSMHFT Lead` = "Informatics",
    `Metric Category Name` = case_when(
      (`Metric Category Type` == "Numerator") ~ "Total number of patients seen within target",
      (`Metric Category Type` == "Denominator") ~ "Total number of patients presented in ED",
      (`Metric Category Type` == "Percentage") ~ "Percentage number of patients seen within target",
    )
  ) %>% 
  select(Metric, `Metric Category Type`, `Metric Category Name`,
         `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)

# Save the data frame in a data list
data_list[["% seen within 1 hour assessment target in ED"]] <- pct_seen_within_target


#3. Delayed discharges -----------------------------------------------------------

metric_categories <- c("Delay Discharges Total",
                       "Adult Acute Total", "Adult Acute - NHS", "Adult Acute - Social Care",
                       "Adult Acute - Both Health and Social Care", "Adult Acute - Housing",
                       "Older Adult Total", " Older Adult - NHS", "Older Adult - Social Care",
                       "Older Adult - Both Health and Social Care", "Older Adult - Housing")


start_row <- which(data_2$Metric == "Delayed Discharges:")[1]
end_row <- which(data_2$Metric == "POS Transfers from ED")[1] - 1

delayed_discharges_dt <- data_2 %>% 
  slice(start_row:end_row) %>% 
  convert_dates_cols() %>% 
  mutate(`BSMHFT Lead` = "Informatics") %>% 
  mutate(Metric = metric_categories) %>% 
  as_tibble() %>% 
  mutate(
    across(.cols = 3:ncol(data_2),
           .fns = ~as.numeric(.))) %>%
  pivot_longer(
    cols = -c('Metric', 'BSMHFT Lead'),
    names_to = "Date",
    values_to = "Value") %>% 
  mutate(
    Value = as.numeric(Value),
    Day = wday(Date, label = TRUE, abbr = FALSE),
    `File Name` = "BSHMFT Daily MH SITREP") %>% 
  mutate(`Metric Category Type` = case_when(
    grepl("Adult Acute", Metric) ~ "Adult Acute",
    grepl("Older Adult", Metric) ~ "Older Adult",
    TRUE ~ "Total" 
  )) %>% 
  mutate(`Metric Category Name` = case_when(
    grepl(" - NHS", Metric) ~ "NHS",
    grepl(" - Social Care", Metric) ~ "Social Care",
    grepl(" - Both Health and Social Care", Metric) ~ "Both Health and Social Care",
    grepl(" - Housing", Metric) ~ "Housing",
    TRUE ~ "Total"
  )) %>% 
  mutate(Metric = "Delayed Discharges",
         Provider = case_when(
           (`Metric Category Type` == "Total" & `Metric Category Name` == "Total") ~ "System",
           TRUE ~ NA
         )) %>% 
  select(Metric, `Metric Category Type`, `Metric Category Name`,
         `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)


# Check count of rows for each metric category
counts <- delayed_discharges_dt %>% 
  group_by(Date) %>% 
  summarise(count = n())

# Save the dataframe in a data list
data_list[["Delayed discharges"]] <- delayed_discharges_dt 

# delayed_discharges_dt %>%
#   filter(Date == "11/04/2022") %>%
#   View()

#4. Discharges -------------------------------------------------------------------

start_row <- which(data_2$Metric == "Discharges:")[1]
end_row <- nrow(data_2)


discharges_dt <- data_2 %>% 
  slice(start_row: end_row) %>% 
  convert_dates_cols() %>% 
  mutate(`BSMHFT Lead` = "Informatics") %>% 
  mutate(Metric = c("Discharges Total", "BSMHFT Total",
                    "Adult Acute - BSMHFT", "Older Adult - BSMHFT",
                    "Out of Area Total", "Adult Acute - Out of Area")) %>%
  as_tibble() %>% 
  mutate(
    across(.cols = 3:ncol(data_2),
           .fns = ~as.numeric(.))) %>%
  pivot_longer(
    cols = -c('Metric', 'BSMHFT Lead'),
    names_to = "Date",
    values_to = "Value") %>% 
  mutate(
    Value = as.numeric(Value),
    Day = wday(Date, label = TRUE, abbr = FALSE),
    Organisation = "BSHMFT",
    `File Name` = "BSMHFT Daily MH SITREP") %>%
  mutate(`Metric Category Type` = case_when(
    grepl("BSMHFT", Metric) ~ "BSHMFT",
    grepl("Out of Area", Metric) ~ "Out of Area",
    TRUE ~ "Total" 
  )) %>% 
  mutate(`Metric Category Name` = case_when(
    grepl("Adult Acute - ", Metric) ~ "Adult Acute",
    grepl("Older Adult - ", Metric) ~ "Older Adult",
    TRUE ~ "Total"
  )) %>% 
  mutate(Metric = "Discharges",
         Provider = case_when(
    (`Metric Category Type` == "Total" & `Metric Category Name` == "Total") ~ "System",
    TRUE ~ NA
  )) %>% 
  select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)

# Check count of rows for each metric category
counts <- discharges_dt %>% 
  group_by(Date) %>% 
  summarise(count = n())

# Save the dataframe in a data list
data_list[["Discharges"]] <- discharges_dt 


# discharges_dt %>%
#   filter(Date == "11/04/2022") %>%
#   View()

#5. Total IP MH that are MFFD ----------------------------------------------------

start_row <- which(data_2$Metric == "Total IP MH that are MFFD:")[1]
end_row <- which(data_2$Metric == "S136 patients in ED:")[1]

total_IP_MH_MFFD <- data_2 %>% 
  slice(start_row:(end_row - 1)) %>% 
  convert_dates_cols() %>% 
  mutate(`BSMHFT Lead` = "Bed Management") %>% 
  mutate(Metric = c("Total IP MH that are MFFD", "City Total",
                    "City - ED", "City - Other wards", "Good Hope Total",
                    "Good Hope - ED", "Good Hope - Other wards",
                    "Heartlands Total", "Heartlands - ED", "Heartlands - Other wards",
                    "QE/UHB Total", "QE/UHB - ED", "QE/UHB - Other wards")) %>%
 as_tibble() %>% 
  mutate(
    across(.cols = 3:ncol(data_2),
           .fns = ~as.numeric(.))) %>%
  pivot_longer(
    cols = -c('Metric', 'BSMHFT Lead'),
    names_to = "Date",
    values_to = "Value") %>% 
  mutate(
    Value = as.numeric(Value),
    Day = wday(Date, label = TRUE, abbr = FALSE),
    `File Name` = "BSMHFT Daily MH SITREP") %>%
  mutate(Provider = case_when(
    grepl("City", Metric) ~ "City",
    grepl("Good Hope", Metric) ~ "Good Hope",
    grepl("Heartlands", Metric) ~ "Heartlands",
    grepl("QE/UHB", Metric) ~ "QE/UHB",
    TRUE ~ "System" )) %>%
  mutate(`Metric Category Type` = case_when(
    grepl(" - ED", Metric) ~ "ED",
    grepl(" - Other wards", Metric) ~ "Other wards",
    TRUE ~ "Total"
  )) %>% 
  mutate(Metric = "Total IP MH that are MFFD",
         `Metric Category Name` = "Total") %>% 
  select(Metric, `Metric Category Type`, `Metric Category Name`,
         `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)

# Check count of rows for each metric category
count <- total_IP_MH_MFFD %>% 
  group_by(Date) %>% 
  summarise(count = n())

# Save the dataframe in a data list
data_list[["Total IP MH that are MFFD"]] <- total_IP_MH_MFFD 

# total_IP_MH_MFFD %>%
#   filter(Date == "11/04/2022") %>%
#   View()

#6. Total IP MH ------------------------------------------------------------------

start_row <- which(data_2$Metric == "Total IP MH:")[1]
end_row <- which(data_2$Metric == "Total IP MH that are MFFD:")[1]

total_IP_MH <- data_2 %>% 
  slice(start_row:(end_row - 1)) %>% 
  convert_dates_cols() %>% 
  mutate(`BSMHFT Lead` = "Bed Management") %>% 
  mutate(Metric = c("Total IP MH", "City Total",
                    "City - ED", "City - Other wards", "Good Hope Total",
                    "Good Hope - ED", "Good Hope - Other wards",
                    "Heartlands Total", "Heartlands - ED", "Heartlands - Other wards",
                    "QE/UHB Total", "QE/UHB - ED", "QE/UHB - Other wards")) %>%
  as_tibble() %>% 
  mutate(
    across(.cols = 3:ncol(data_2),
           .fns = ~as.numeric(.))) %>%
  pivot_longer(
    cols = -c('Metric', 'BSMHFT Lead'),
    names_to = "Date",
    values_to = "Value") %>% 
  mutate(
    Value = as.numeric(Value),
    Day = wday(Date, label = TRUE, abbr = FALSE),
    `File Name` = "BSMHFT Daily MH SITREP") %>%
  mutate(Provider = case_when(
    grepl("City", Metric) ~ "City",
    grepl("Good Hope", Metric) ~ "Good Hope",
    grepl("Heartlands", Metric) ~ "Heartlands",
    grepl("QE/UHB", Metric) ~ "QE/UHB",
    TRUE ~ "System" )) %>%
  mutate(`Metric Category Type` = case_when(
    grepl(" - ED", Metric) ~ "ED",
    grepl(" - Other wards", Metric) ~ "Other wards",
    TRUE ~ "Total"
  )) %>% 
  mutate(Metric = "Total IP MH",
         `Metric Category Name` = NA) %>% 
  select(Metric, `Metric Category Type`, `Metric Category Name`,
         `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)

# Check count of rows for each metric category

counts <- total_IP_MH %>% 
  group_by(Date) %>% 
  summarise(count = n())

# Save the dataframe in a data list
data_list[["Total IP MH"]] <- total_IP_MH 

# total_IP_MH %>%
#   filter(Date == "11/04/2022") %>%
#   View()

# 7. Write output --------------------------------------------------------------

# Combine all data frames into one
all_data <- bind_rows(data_list, .id = "Metric")

# Add the combined data frame to the list of data under the key "All"
data_list$All <- all_data

# Update Date column 
data_list <- lapply(data_list, function(df) {
  if("Date" %in% names(df)) {
    # Ensure the Date column is a Date object
    if(!inherits(df$Date, "Date")) {
      df$Date <- dmy(df$Date) 
    }
    
    # Directly modify the year for dates in 2032
    is_2032 <- year(df$Date) == 2032
    df$Date[is_2032] <- update(df$Date[is_2032], year = 2023)
  }
  return(df)
})

# Write each processed data frame to a new sheet in an Excel workbook
write_xlsx(data_list, path = output_file_name)

# data_list[["MFFD Transferred Out"]] %>% 
#   filter(Date == "2032-09-18") %>% 
#   View()

#8. Move the raw data to "Old data" folder once data processing is finished-----
## Define the destination path explicitly as a subdirectory of `file_path`
destination_path <- paste0(file_path, "/Old data")


# Ensure the destination directory exists; create it if it doesn't
if (!dir.exists(destination_path)) {
  dir.create(destination_path)
}

# List the Excel file(s) we intend to move
excel_files <- list.files(path = file_path, pattern = "\\.xlsx$", full.names = TRUE)

# Check if there are Excel files to move
if (length(excel_files) > 0) {
  # Since there's always one file, directly move it
  file_move(excel_files[1], destination_path)
} else {
  # If no Excel files are found, print a message
  message("No files found to move.")
}
