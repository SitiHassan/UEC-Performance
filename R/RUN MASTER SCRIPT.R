library(readxl)
library(readr)
library(dplyr)
library(janitor)
library(tidyr)
library(lubridate)
library(stringr)
library(writexl)
library(fs)

# Clear the work space, excluding libraries, parameters and the run_master_script function
rm(list = setdiff(ls(), c("run_master_script", "directory_path", "input_file_path", "output_file_path", lsf.str())))

# The path to the parent work directory
directory_path = "./Reports/01_Adhoc/BSOL_1339_UEC_Performance"

# The path to the input excel files
input_file_path <- paste0(directory_path, "/Raw Data/Recent data")
  
# The path the output excel files
output_file_path <- paste0(directory_path, "/Processed Data/Recent data")


run_master_script <- function(){
  
  
  start_time = Sys.time()
  
  print("Script 1: Starting to extract ADA data..")
  source(paste0(directory_path, "R scripts/01_01_Extract_ADA_Report.R"))
  print("Finished extracting ADA data.")
  
  print("Script 2: Starting to extract BSMHFT Sitrep")
  source(paste0(directory_path, "R scripts/02_Extract_BSMHFT_Daily_MH_Sitrep.R"))
  print("Finished extracting BSMHFT Sitrep.")
  
  print("Script 3: Starting to extract BWC Sitrep")
  source(paste0(directory_path, "R scripts/03_Extract_BWC_Weekly_Sitrep.R"))
  print("Finished extracting BWC Sitrep.")
  
  print("Script 4: Starting to extract UHB Sitrep")
  source(paste0(directory_path, "R scripts/04_Extract_UHB_Daily_Sitrep.R"))
  print("Finished extracting UHB Sitrep.")
  
  print("Script 5: Starting to extract Virtual Wards data")
  source(paste0(directory_path, "R scripts/05_Extract_Virtual_Wards.R"))
  print("Finished extracting Virtual Wards data.")
  
  print("Script 6: Starting to extract WMAS data")
  source(paste0(directory_path, "R scripts/06_Extract_WMAS.R"))
  print("Finished extracting VMAS data.")
  
  print("Script 7: Starting to extract UHB Discharges data")
  source(paste0(directory_path, "R scripts/07_Extract_UHB_Discharges.R"))
  print("Finished extracting UHB Discharges data.")
  
  print("Script :8 Starting to extract BCHC UCR data")
  source(paste0(directory_path, "R scripts/08_Extract_BCHC_UCR.R"))
  print("Finished extracting BCHC UCR data.")
  
  end_time = Sys.time()
  
  time_taken = end_time - start_time
  
  print(paste0("Total time taken to run all scripts: ", time_taken))
  
}

run_master_script()