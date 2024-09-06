# Path to the Excel file
file_path <- paste0(input_file_path, "/UHB Discharges")

# Get a list of all Excel files in the folder (.xlsx and .xls)
excel_files <- list.files(path = file_path, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_UHB_Discharges.xlsx")


process_data <- function(excel_file){
  data <- read_xlsx(excel_file[1])
  
  data %>% 
    mutate(Metric = "Daily Pathway Discharges",
           `Metric Category Name` = "Total",
           `Provider` = "UHB",
           Date = as.Date(Date),
           Day = wday(Date, label = TRUE, abbr = FALSE),
           Value = as.numeric(Value)) %>% 
    rename(`Metric Category Type` = Pathway) %>% 
    mutate(`File Name` = "UHB Daily Pathway") %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`, Provider, Date, Day, Value, `File Name`)
  
}

# Apply function to to a list of Excel files in the WMAS folder
list_of_processed_dfs <- lapply(excel_files, process_data)

# Combine all processed data frames into one data frame
output <- bind_rows(list_of_processed_dfs) %>% 
  dplyr::arrange(Date) %>%
  clean_names(case = "title")


# Write the processed data frame in an Excel workbook
write_xlsx(output, path = output_file_name) 

# # Move the raw data to "Old data" folder once data processing is finished
# ## Define the destination path explicitly as a subdirectory of `file_path`
# destination_path <- paste0(file_path, "/Old data")
# 
# # Ensure the destination directory exists; create it if it doesn't
# if (!dir.exists(destination_path)) {
#   dir.create(destination_path)
# }
# # List the Excel file(s) we intend to move
# excel_files <- list.files(path = file_path, pattern = "\\.xlsx$", full.names = TRUE)
# 
# # Check if there are Excel files to move
# if (length(excel_files) > 0) {
#   # Since there's always one file, directly move it
#   file_move(excel_files[1], destination_path)
# } else {
#   # If no Excel files are found, print a message
#   message("No files found to move.")
# }