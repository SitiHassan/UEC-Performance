
# Data Processing --------------------------------------------------------------

# The function logics:
#1. Make the row containing "Metric" as header
#2. Select the necessary columns, excluding the column with NAs & the 1st column
#3. Rename the column "Description" to "Metric"
#4. Create a tibble
#5. Pivot the wider dates format to a long format
#6. Format the dates to appropriate "DD/MM/YYYY" format
#7. Create a new column "Day" for the corresponding day
#8. Create a new column "Organisation" to distinguish among trusts
#9. Ensure all columns are in the correct data types
#10.Re-arrange the columns for the final output


process_data <- function(excel_file){
  
  data <- read_xlsx(excel_file[1])
   
  data %>% 
    janitor::row_to_names(row_number = which(data$`BWC Informatics` == "Metric")) %>%
    select(Description, c(4: ncol(data))) %>%
    rename(Metric = Description) %>% 
    as_tibble() %>% 
    pivot_longer(
      cols = -Metric,
      names_to = "Date",
      values_to = "Value"
    ) %>% 
    mutate(
      Date = dmy(Date),
      Day = lubridate::wday(Date, label = TRUE, abbr = FALSE),
      Provider = "BWC",
      `Metric Category Type` = "Total",
      `Metric Category Name` = "Total",
      `File Name` = "BWC Weekly SITREP Report"
    ) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           Provider, Date, Day, Value, `File Name`)
}



# Path to the Excel file
file_path <- paste0(input_file_path, "/BWC Sitrep")

# Access the Excel file within the BWC Sitrep folder
excel_file <- list.files(path = file_path, full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_BWC_Sitrep.xlsx")

output <- process_data(excel_file)

# Write the processed data frame to a new sheet in an Excel workbook
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