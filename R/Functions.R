# install.packages("usethis")
# install.packages("devtools")
# install.packages("roxygen2")
# usethis::create_package(".")

# devtools::document() # To generate documentation
# devtools::load_all() # To load the package locally
# devtools::build()    # To build packages



#' List the most recent files in a SharePoint folder
#'
#' This function connects to a SharePoint folder, lists all files, and returns the most recent X files,
#' sorted by the date embedded in their filenames.
#'
#' The filenames are assumed to follow the format "YYYY-MM-DD" for date extraction. Files without valid
#' date formats are ignored.
#'
#' @param X An integer specifying the number of most recent files to return. Default is 14.
#' @param folder_name A character string specifying the SharePoint folder name to search within.
#' @param site_url A character string containing the SharePoint site URL.
#'
#' @return A data frame containing the most recent X files with columns for file name, size, ID, and extracted date.
#' @export
#'
#' @examples
#' \dontrun{
#'   # List the most recent 14 files in the specified SharePoint folder
#'   recent_files <- list_recent_spt_files(
#'     X = 14,
#'     folder_name = "Daily UEC Performance/Raw data/Recent data/ADA",
#'     site_url = "https://csucloudservices.sharepoint.com/sites/BSOLEmbeddedBI"
#'   )
#'   print(recent_files)
#' }

list_recent_spt_files <- function(X = 14, folder_name, site_url) {

  # Access the SharePoint site
  CSU_site <- get_sharepoint_site(
    site_url = site_url
  )

  # Access the default document library
  doc_library <- CSU_site$get_drive()

  # Navigate to the data folder using the provided folder name
  data_folder <- doc_library$get_item(folder_name)

  # List files in the folder (as a data frame)
  files <- data_folder$list_items()

  # Check if `files` contains data
  if (nrow(files) == 0) {
    stop("No files found in the folder or the folder path is incorrect.")
  }

  # Extract dates from the filenames (assuming the format "YYYY-MM-DD" appears in the name)
  files_with_dates <- data.frame(
    name = files$name,
    size = files$size,
    id = files$id,
    extracted_date = as.POSIXct(sub("([0-9]{4}-[0-9]{2}-[0-9]{2})T.*", "\\1", files$name), format="%Y-%m-%d")
  )

  # Remove entries where the date could not be extracted (NA values)
  files_with_dates <- files_with_dates[!is.na(files_with_dates$extracted_date), ]

  # Sort files by the extracted date in descending order
  files_sorted <- files_with_dates[order(files_with_dates$extracted_date, decreasing = TRUE), ]

  # Return only the top X filenames (default 14)
  return(head(files_sorted, X))
}

#' Download the latest files from a SharePoint folder
#'
#' This function connects to a specified SharePoint folder and downloads the most recent files
#' to a specified local directory. The files are downloaded based on the file names provided in the
#' `recent_files` data frame, which should contain the names of the files to be downloaded.
#'
#' @param recent_files A data frame containing the file names to be downloaded. Typically, this would be
#'        generated using a function like `list_recent_spt_files()`.
#' @param folder_name A character string specifying the SharePoint folder name from which to download the files.
#' @param directory A character string specifying the local directory where the files will be saved.
#' @param site_url A character string containing the SharePoint site URL.
#'
#' @return Downloads the specified files to the local directory. Returns nothing (NULL).
#' @export
#'
#' @examples
#' \dontrun{
#'   # List and download the most recent files from the specified SharePoint folder
#'   recent_files <- list_recent_spt_files(
#'     X = 14,
#'     folder_name = "Daily UEC Performance/Raw data/Recent data/ADA",
#'     site_url = "https://csucloudservices.sharepoint.com/sites/BSOLEmbeddedBI"
#'   )
#'   download_spt_files(
#'     recent_files = recent_files,
#'     folder_name = "Daily UEC Performance/Raw data/Recent data/ADA",
#'     directory = paste0(directory, "/Data/New ADA"),
#'     site_url = "https://csucloudservices.sharepoint.com/sites/BSOLEmbeddedBI"
#'   )
#' }
download_spt_files <- function(recent_files, folder_name, directory, site_url) {

  # Access the SharePoint site
  CSU_site <- get_sharepoint_site(
    site_url = site_url
  )

  # Access the default document library
  doc_library <- CSU_site$get_drive()

  # Navigate to the data folder using the provided folder name
  data_folder <- doc_library$get_item(folder_name)

  for (i in 1:nrow(recent_files)) {

    # Extract the file name
    file_name <- recent_files$name[i]

    # Access the file in SharePoint
    sharepoint_file <- data_folder$get_item(file_name)

    # Download the file to the specified local directory, overwriting any existing files
    sharepoint_file$download(file.path(directory, file_name), overwrite = TRUE)
  }
}


#' Return the most recent Excel files from a directory
#'
#' This function lists Excel files from a specified directory and returns the most recent X files,
#' based on their modification time. It supports both `.xlsx` and `.xls` file formats.
#'
#' @param directory A character string specifying the path to the directory containing the Excel files.
#' @param X An integer specifying the number of most recent files to return. Default is 14.
#'
#' @return A character vector containing the filenames (with full paths) of the most recent X Excel files.
#' @export
#'
#' @examples
#' \dontrun{
#'   # Get the 14 most recent Excel files from the directory
#'   latest_filenames <- get_latest_excel_filenames(
#'     directory = paste0(directory, "/Data/New ADA"),
#'     X = 14
#'   )
#'   print(latest_filenames)
#' }
get_latest_excel_filenames <- function(directory, X = 14) {

  # Get a list of all Excel files in the folder (.xlsx and .xls)
  excel_files <- list.files(path = directory, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

  # Check if there are any files in the directory
  if (length(excel_files) == 0) {
    stop("No Excel files found in the directory.")
  }

  # Get file modification times
  file_info <- file.info(excel_files)

  # Sort files by modification time in descending order (latest first)
  sorted_files <- rownames(file_info[order(file_info$mtime, decreasing = TRUE), ])

  # Get the latest X files
  latest_files <- head(sorted_files, X)

  return(latest_files)  # Return just the filenames
}

#' Read the most recent Excel files into data frames
#'
#' This function takes a list of file paths to Excel files and reads each file into a data frame.
#' The resulting data frames are returned in a list. The first row of each Excel file is not used as
#' column names, and all columns are read without headers.
#'
#' The `latest_files` argument, which is a list of file paths, can be obtained using the
#' `get_latest_excel_filenames()` function to first retrieve the most recent Excel files from a directory.
#'
#' @param latest_files A character vector containing the full paths to the Excel files to be read.
#'
#' @return A list of data frames, where each data frame corresponds to an Excel file from `latest_files`.
#' @export
#'
#' @examples
#' \dontrun{
#'   # Get the latest filenames from the directory
#'   latest_filenames <- get_latest_excel_filenames(
#'     directory = "path/to/excel/files",
#'     X = 14
#'   )
#'   # Read the Excel files into data frames
#'   latest_dataframes <- get_latest_excel_dataframes(latest_filenames)
#'   # Access the first data frame
#'   print(latest_dataframes[[1]])
#' }
get_latest_excel_dataframes <- function(latest_files) {

  # Read each Excel file into a data frame and store in a list (without using the first row as column names)
  list_of_dataframes <- lapply(latest_files, read_excel, col_names = FALSE)

  return(list_of_dataframes)  # Return the list of data frames
}
#' Process and combine ADA data frames
#'
#' This function processes a list of ADA data frames and combines the processed results into a single data frame.
#' It extracts the date from the file name, processes the data (dropping NA values and converting necessary columns
#' to numeric), and combines the results into a single table, filtering by a provided date.
#'
#' The `latest_dataframes` argument, which is a list of data frames, can be obtained using the
#' `get_latest_excel_dataframes()` function. The `latest_filenames` argument, which is a vector of file names,
#' can be obtained using the `get_latest_excel_filenames()` function.
#'
#' @param latest_dataframes A list of data frames containing ADA data to be processed. Can be obtained using `get_latest_excel_dataframes()`.
#' @param latest_filenames A character vector of file names corresponding to the ADA data frames. Can be obtained using `get_latest_excel_filenames()`.
#' @param latest_date A Date object representing the cutoff date to filter the combined data.
#' @param drop_na A logical value indicating whether to drop rows with missing values in each data frame. Default is TRUE.
#'
#' @return A combined data frame containing processed ADA data, filtered by the specified date.
#' @export
#'
#' @examples
#' \dontrun{
#'   # Obtain the latest filenames and data frames
#'   latest_filenames <- get_latest_excel_filenames(directory = "path/to/excel/files", X = 14)
#'   latest_dataframes <- get_latest_excel_dataframes(latest_filenames)
#'
#'   # Process and combine the ADA data
#'   combined_df <- process_ADA_data(
#'     latest_dataframes = latest_dataframes,
#'     latest_filenames = latest_filenames,
#'     latest_date = as.Date("2024-09-01")
#'   )
#'   print(combined_df)
#' }

process_ADA_data <- function(latest_dataframes, latest_filenames, latest_date, drop_na = TRUE) {

  # Internal function to process each ADA data frame
  process_ADA <- function(df, header_row_value = 'Site', drop_na = TRUE, file_name = "unknown file") {

    # Extract date from the file name and subtract 1 day to get the true date
    date_from_filename <- as.Date(sub("([0-9]{4}-[0-9]{2}-[0-9]{2})T.*", "\\1", basename(file_name)), format = "%Y-%m-%d")
    true_date <- date_from_filename - 1  # Subtract 1 day as the file contains data from the previous day

    # Detect the row with the header (e.g., 'Site')
    site_row_index <- which(df[[1]] == header_row_value)

    if (length(site_row_index) == 0) {
      stop(paste("Header row not found for file:", file_name))
    }

    # Set column names using the row that contains the header
    names(df) <- as.character(df[site_row_index, ])

    # Remove the rows above the header row (including the header row itself)
    df <- df[(site_row_index + 1):nrow(df), ]

    # Drop rows with NA values if requested
    if (drop_na) {
      df <- df %>% drop_na()
    }

    # Add error handling for numeric conversion
    df <- df %>%
      mutate(Date = true_date) %>%
      mutate(
        `Patients In ADA` = suppressWarnings(as.numeric(`Patients In ADA`)),
        `Patients in ADA Over 4 Hours` = suppressWarnings(as.numeric(`Patients in ADA Over 4 Hours`))
      ) %>%
      rename(Patients_In_ADA = `Patients In ADA`,
             Patients_In_ADA_Over_4_Hrs = `Patients in ADA Over 4 Hours`) %>%
      select(Date, Site, Patients_In_ADA, Patients_In_ADA_Over_4_Hrs)

    return(df)
  }

  # Process each data frame along with the corresponding file name
  list_of_processed_dfs <- lapply(seq_along(latest_dataframes), function(i) {
    df <- latest_dataframes[[i]]
    file_name <- latest_filenames[i]

    tryCatch({
      process_ADA(df, file_name = file_name, drop_na = drop_na)
    }, error = function(e) {
      warning(paste("Error processing file:", file_name, "-", e))
      return(NULL)  # Skip the file with an error
    })
  })

  # Combine the data from all processed data frames
  combined_df <- bind_rows(list_of_processed_dfs) %>%
    filter(!(is.na(Date))) %>%
    filter(Date > latest_date) %>%
    dplyr::arrange(Date)

  return(combined_df)
}


#' Move all files from one folder to another within a main directory
#'
#' This function moves all files from a specified source folder to a destination folder within the
#' main directory. If the destination folder does not exist, it is created. The function overwrites
#' any files in the destination folder if they already exist.
#'
#' @param main_directory A character string specifying the main directory that contains both the source
#'        and destination folders.
#' @param source_folder A character string specifying the name of the source folder from which files will be moved.
#' @param destination_folder A character string specifying the name of the destination folder to which files will be moved.
#'
#' @return Returns `NULL` after moving all files from the source to the destination folder.
#'         Prints a success message for each moved file.
#' @export
#'
#' @examples
#' \dontrun{
#'   # Move all files from 'New ADA' folder to 'Old ADA' folder within the main directory
#'   move_all_files(
#'     main_directory = file.path(directory, "Data"),
#'     source_folder = "New Data",
#'     destination_folder = "Old Data"
#'   )
#' }
move_all_files <- function(main_directory, source_folder, destination_folder) {

  # Define the full paths for the source and destination folders
  source_folder_path <- file.path(main_directory, source_folder)
  destination_folder_path <- file.path(main_directory, destination_folder)

  # Ensure the destination folder exists; if not, create it
  if (!dir_exists(destination_folder_path)) {
    dir_create(destination_folder_path)
  }

  # List all files in the source folder
  files_to_move <- dir_ls(source_folder_path, recurse = FALSE)  # List files, no subdirectories

  # Check if there are any files to move
  if (length(files_to_move) == 0) {
    cat("No files found in the source folder to move.\n")
    return(NULL)  # Exit if no files are found
  }

  # Loop through each file and move it to the destination folder
  for (file_path in files_to_move) {

    # Extract the file name from the full path
    file_name <- basename(file_path)

    # Define the destination path
    destination_file <- file.path(destination_folder_path, file_name)

    # Move the file to the destination directory (overwrite if it already exists)
    file_move(file_path, destination_file)

    # Print success message
    cat("Moved file:", file_name, "to", destination_folder, "directory.\n")
  }
}

