
directory <- "./BSOL 1339 UEC Performance/BCHC UCR"

# Get a list of all Excel files in the One Drive folder (.xlsx and .xls)
excel_files <- list.files(path = paste0(directory, "/New data"), pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

# Check if any files exist
if (length(excel_files) > 0) {
  # Get the modification times of the files
  modification_times <- file.info(excel_files)$mtime
  
  # Find the index of the latest modified file
  latest_index <- which.max(modification_times)
  
  # Get the path of the latest file
  latest_file <- excel_files[latest_index]
  
  # Print the path of the latest file
  cat("Latest file:", latest_file, "\n")
  
  destination_path <- paste0(input_file_path, "/BCHC UCR")
  new_file_path <- file.path(destination_path, basename(latest_file))
  
  # Copy the file to the new location
  file.copy(latest_file, new_file_path)
  
} else {
  message("No Excel files found in the folder.")
}


# Path to the Excel file
file_path <- paste0(input_file_path, "/BCHC UCR")

# Get a list of all Excel files in the BCHC UCR folder (.xlsx and .xls)
excel_files <- list.files(path = file_path, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_BCHC_UCR.xlsx")

data <- read_xlsx(excel_files[1], sheet = "Detail")

# Define function to process data
process_data <- function(data, WMAS_filter = NULL, breach_filter = NULL, team_filter = NULL, metric){
  
  if (!is.null(WMAS_filter)){
    data <- filter(data, WMAS == WMAS_filter)
  }
  if (!is.null(team_filter)){
    data <- filter(data, Team == team_filter)
  }
  if (!is.null(breach_filter)){
    data <- filter(data, Breach == breach_filter)
  }
  
  data %>% 
    group_by(`Referral Date`) %>% 
    summarise(Value = n()) %>% 
    mutate(Metric = metric,
           `Metric Category Type` = "Total",
           `Metric Category Name` = "Total",
           Provider = "BCHC") %>% 
    rename(Date = `Referral Date`) %>% 
    mutate(Date = as.Date(Date),
           Day = wday(Date, label = TRUE, abbr = FALSE),
           `File Name` = "BCHC UCR") %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`, Provider, Date, Day, Value, `File Name`)
}

# Total UCR Referrals
total_ucr_referrals <- process_data(data, metric = "Total UCR referrals")
total_ucr_referrals_without_breaches <- process_data(data, breach_filter = "0", 
                                                     metric = "Total UCR referrals without breaches") 
total_ucr_referrals_with_breaches <- process_data(data, breach_filter = "1",
                                                  metric= "Total UCR referrals with breaches") 

# District Nursing Team(s)
DN_team <- process_data(data, team_filter = "DN", 
                        metric= "Total DN team referrals") 
DN_team_without_breaches <- process_data(data, team_filter = "DN", breach_filter = "0", 
                                         metric = "Total DN team referrals without breaches") 
DN_team_with_breaches <- process_data(data, team_filter = "DN", breach_filter = "1",
                                      metric = "Total DN team referrals with breaches") 

# UCR Team
ucr_team  <- process_data(data, team_filter = "UCR", 
                          metric = "Total UCR team referrals")
ucr_team_without_breaches <- process_data(data, team_filter = "UCR", breach_filter = "0",
                                          metric = "Total UCR team referrals without breaches")
ucr_team_with_breaches <- process_data(data, team_filter = "UCR", breach_filter = "1",
                                       metric = "Total UCR team referrals with breaches")

# Total UCR Referrals via WMAS
total_ucr_referrals_via_WMAS <- process_data(data, WMAS_filter = "1",
                                             metric = "Total UCR referrals via WMAS")
total_ucr_referrals_via_WMAS_without_breaches <- process_data(data, WMAS_filter = "1", breach_filter = "0",
                                                              metric = "Total UCR referrals via WMAS without breaches")
total_ucr_referrals_via_WMAS_with_breaches <- process_data(data, WMAS_filter = "1", breach_filter = "1",
                                                           metric = "Total UCR referrals via WMAS with breaches")

#2. Calculate 2hr performance = referrals within 2hrs (without breaches) / total referrals
calculate_2hr_performance <- function(numerator_data, denominator_data, numerator_metric, denominator_metric){
  combined_data <- bind_rows(numerator_data, denominator_data)
  
  performance_data <- combined_data %>% 
    pivot_wider(
      names_from = Metric,
      values_from = Value
    ) %>% 
  mutate(
    `2hr Performance` = !!sym(numerator_metric) / !!sym(denominator_metric)
  ) %>% 
    pivot_longer(
      cols = c(7:9),
      names_to = "Metric",
      values_to = "Value"
    ) %>%
    mutate(`Metric Category Type` = case_when(
      Metric == "2hr Performance" ~ denominator_metric,
      TRUE ~ "Total")) %>% 
    mutate(`Metric Category Name` = case_when(
      Metric == "2hr Performance" ~ "Percentage",
      TRUE ~ "Total")) %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`, Provider, Date, Day, Value, `File Name`)
  
  return(performance_data)
  
}

total_ucr_referrals_performance <- calculate_2hr_performance(numerator_data = total_ucr_referrals_without_breaches,
                                                             denominator_data = total_ucr_referrals,
                                                             numerator_metric = "Total UCR referrals without breaches",
                                                             denominator_metric = "Total UCR referrals")
  
DN_team_performance <- calculate_2hr_performance(numerator_data = DN_team_without_breaches,
                                                            denominator_data = DN_team,
                                                            numerator_metric = "Total DN team referrals without breaches",
                                                            denominator_metric = "Total DN team referrals")

UCR_team_performance <- calculate_2hr_performance(numerator_data = ucr_team_without_breaches,
                                                        denominator_data = ucr_team,
                                                        numerator_metric = "Total UCR team referrals without breaches",
                                                        denominator_metric = "Total UCR team referrals")
  
total_UCR_referrals_via_WMAS_performance <- calculate_2hr_performance(numerator_data = total_ucr_referrals_via_WMAS_without_breaches,
                                                        denominator_data = total_ucr_referrals_via_WMAS,
                                                        numerator_metric = "Total UCR referrals via WMAS without breaches",
                                                        denominator_metric = "Total UCR referrals via WMAS")

# Combine all data
output <- bind_rows(total_ucr_referrals_performance, total_ucr_referrals_with_breaches,
                    DN_team_performance, DN_team_with_breaches,
                    UCR_team_performance, ucr_team_with_breaches,
                    total_UCR_referrals_via_WMAS_performance, total_ucr_referrals_via_WMAS_with_breaches)


output <- output %>% 
  dplyr::arrange(Date) %>%
  clean_names(case = "title")

# Write the processed data frame in an Excel workbook
write_xlsx(output, path = output_file_name)

# Move the raw data to "Old data" folder once data processing is finished
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