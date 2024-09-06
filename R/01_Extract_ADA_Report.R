max_date <- "2024-04-29"

directory <- "./BSOL 1339 UEC Performance/ADA data"

# Get a list of all Excel files in the folder (.xlsx and .xls)
excel_files <- list.files(path = paste0(directory, "/New data"), pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

# Copy these last 7 excel files and save them in the ADA folder
# Check if any files exist
if (length(excel_files) > 0) {
  # Get the modification times of the files
  modification_times <- file.info(excel_files)$mtime
  
  # Find the index of the latest modified 7 files
  latest_indexes <- order(modification_times, decreasing = TRUE)[1:13]
  
  # Get the path of the latest files
  latest_files <- excel_files[latest_indexes]
  
  # Print the path of the latest file
  cat("Latest files:", latest_files, "\n")
  
  destination_path <- paste0(input_file_path, "/ADA/new data")
  
  # Create the destination directory if it doesn't exist
  if (!dir.exists(destination_path)) {
    dir.create(destination_path, recursive = TRUE)
  }
  
  for (file in latest_files){
    new_file_path <- file.path(destination_path, basename(file))
    file.copy(file, new_file_path)
  }
  
  
} else {
  message("No Excel files found in the folder.")
}

# Read each Excel file into a data frame and store in a list
list_of_dataframes <- lapply(excel_files, read_excel)

# Function to process each dataframe
process_dataframe <- function(df) {
  
  # Extract date from the first column (assumes date is always in first row and column)
  date_string <- names(df)[1]
  date_extracted <- as.Date(sub(".*on\\s([0-9]{2}/[0-9]{2}/[0-9]{4}).*", "\\1", date_string),
                            format = "%d/%m/%Y")
  
  # Find the row with the string 'Site' and use it as the header
  site_row_index <- which(df[[1]] == 'Site')
  names(df) <- as.character(df[site_row_index, ])
  df <- df[(site_row_index + 1):nrow(df), ]
  
  # Remove rows with any NA values
  df <- df %>% 
    drop_na() %>% 
    mutate(Date = date_extracted,
           Day = lubridate::wday(Date, label = TRUE, abbr = FALSE)) %>% 
    mutate(
      `Patients In ADA` = as.numeric(`Patients In ADA`),
      `Patients in ADA Over 4 Hours` = as.numeric(`Patients in ADA Over 4 Hours`)) %>% 
    rename(Provider = Site) %>%
    pivot_longer(
      cols = c("Patients In ADA", "Patients in ADA Over 4 Hours"),
      names_to = "Metric",
      values_to = "Value",
    ) %>% 
    mutate(
      `Metric Category Type` = "Total",
      `Metric Category Name` = "Total",
      `File Name` = "Daily ADA Report"
    ) %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           Provider, Date, Day, Value, `File Name`)
  
  
  return(df)
}

# Apply the custom function to each df in a list of dfs
list_of_processed_dfs <- lapply(list_of_dataframes, process_dataframe)

# Combine the data from all dfs into a main table
output <- bind_rows(list_of_processed_dfs) %>% 
  dplyr::arrange(Date) %>%
  clean_names(case = "title",
              abbreviations = c("ADA")) %>% 
  filter(!(is.na(Date)) &
           Date > max_date)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_ADA.xlsx")


# Write the processed data frame to a new sheet in an Excel workbook
write_xlsx(output, path = output_file_name)

# Move the raw data to "Old data" folder once data processing is finished
## Define the destination path explicitly as a subdirectory of `file_path`
destination_path <- paste0(input_file_path, "/ADA/Old data")

# Ensure the destination directory exists; create it if it doesn't
if (!dir.exists(destination_path)) {
  dir.create(destination_path)
}

# List the Excel file(s) we intend to move
excel_files <- list.files(path = paste0(input_file_path, "/ADA"), pattern = "\\.xlsx$", full.names = TRUE)


# Check if there are Excel files to move
if (length(excel_files) > 0) {
  for (file in excel_files) {

    # Move the file to "Old data" folder
    file_move(file, destination_path)
  }
} else {
  message("No files found to move.")
}

