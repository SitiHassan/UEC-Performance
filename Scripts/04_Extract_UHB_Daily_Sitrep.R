
# Data Processing --------------------------------------------------------------

# The function logics:
#1. Remove first 6 rows
#2. Make the first row as headers/columns
#3. Rename the first column to "Metric"
#4. Create a tibble
#5. Pivot the wider dates format to a long format
#6. Format the Excel dates to appropriate "DD/MM/YYYY" format
#7. Create a new column "Day" for the corresponding day
#8. Create a new column "Organisation" to distinguish among trusts
#9. Ensure all columns are in the correct data types
#10.Re-arrange the columns for the final output


process_sheet_data <- function(excel_file, sheet_name) {
  read_xlsx(excel_file[1], sheet = sheet_name) %>%
    slice(-1:-6) %>%
    janitor::row_to_names(row_number = 1) %>%
    rename(Metric = 1) %>%
    as_tibble() %>%
    pivot_longer(cols = -Metric, names_to = "Date", values_to = "Value") %>%
    mutate(Date = as.Date(as.numeric(Date), origin = "1899-12-30")) %>%
    mutate(Day = lubridate::wday(Date, label = TRUE, abbr = FALSE)) %>%
    mutate(Provider = sheet_name) %>%
    mutate(Value = str_replace_all(Value, "%", "")) %>%
    mutate(Value = readr::parse_number(Value)) %>%
    mutate(`File Name` = "Daily SITREP Checksheet") %>% 
    mutate(`Metric Category Type` = "Total",
           `Metric Category Name` = "Total") %>% 
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           Provider, Date, Day, Value, `File Name`)
}

# Path to the Excel file
file_path <- paste0(input_file_path, "/UHB Sitrep")

# Access the Excel file within the UHB Sitrep folder
excel_file <- list.files(path = file_path, full.names = TRUE)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_UHB_Sitrep.xlsx")


# Get the names of all sheets in the Excel file
sheet_names <- excel_sheets(excel_file[1])

# Initialize an empty list to store data frames
data_list <- list()

# Loop through each sheet name and process the data
for(sheet_name in sheet_names) {
  data_list[[sheet_name]] <- process_sheet_data(excel_file, sheet_name)
}

# Combine all data frames into one
all_data <- bind_rows(data_list, .id = "Provider")

# Add the combined data frame to the list of data under the key "All"
data_list$All <- all_data

# Write each processed data frame to a new sheet in an Excel workbook
write_xlsx(data_list, path = output_file_name)

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

# ------------------------------------------------------------------------------




