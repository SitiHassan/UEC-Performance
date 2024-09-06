
process_data <- function(excel_file) {
  
  data <- read_xlsx(excel_file[1])

  date <- colnames(data)[3]
  
  data %>% 
    janitor::row_to_names(row_number = 1) %>%
    as_tibble() %>% 
    mutate(Date = date) %>% 
    mutate(Date = as.Date(as.numeric(Date), origin = "1899-12-30"),
           Day = wday(Date, label = TRUE, abbr = FALSE)) %>% 
    mutate(`File Name` = "Virtual Ward") %>% 
    rename(Percentage = `%`) %>% 
    pivot_longer(
      cols = c("Capacity", "Occupancy", "Percentage"),
      names_to = "Metric",
      values_to = "Value"
    ) %>% 
    mutate(Value = parse_number(Value)) %>% 
    rename(`Metric Category Type` = `Virtual Ward`) %>% 
    mutate(`Metric Category Type` = str_replace_all(`Metric Category Type`, "\\*", "")) %>% 
    mutate(`Metric Category Name` = "Total") %>% 
  mutate(Provider = case_when(
    str_detect(`Metric Category Type`, "Bham only") ~ "UHB",
    str_detect(`Metric Category Type`, "QEHB") ~ "Queen Elizabeth",
    TRUE ~ NA_character_)) %>% 
    mutate(`Metric Category Type` = str_replace(`Metric Category Type`, " \\(Bham only\\)", "")) %>% 
    mutate(`Metric Category Type` = str_replace(`Metric Category Type`, "QEHB @ ", "")) %>% 
    select(`Metric`, `Metric Category Type`, `Metric Category Name`, Provider,
           Date, Day, Value, `File Name`) 
}

# Path to the Excel file
file_path <- paste0(input_file_path, "/Virtual Wards")

# Access the Excel file within the Virtual Wards folder
excel_file <- list.files(path = file_path, pattern = "\\.xlsx$|\\.xls$",full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_Virtual_Wards.xlsx")


output <- process_data(excel_file)

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