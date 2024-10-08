# UECETL

## Overview
**UECETL** is an ETL (Extract, Transform, Load) tool designed to process multiple sources of datasets for analyzing **Urgent and Emergency Care (UEC)** performance. The tool is currently in development and focuses on compiling functions that streamline the analysis of UEC performance metrics. It extracts daily raw local submissions of performance data and processes them for further analysis.

## Features
- **ETL Functions**: Extract and transform data from various sources.
- **Data Processing**: Handle daily raw local submissions of UEC performance metrics.
- **Performance Metrics**: Supports urgent and emergency care performance tracking.
- **Customizable**: Functions can be easily adapted for different data sources.

## Installation
This package is still in development and can be installed locally by cloning this repository and using `devtools`:

```r
# Clone the repository
# Install the development version of the package
devtools::install()

```

## How to Use
Once installed, the **UECETL** package provides a series of functions to extract, process, and anlyze UEC Performance data.

```r
# Load the package
library(UECETL)

# Example of processing a dataset
processed_data <- process_ADA_data(latest_dataframes, latest_filenames, latest_date)

```

This package is **in development** and currently focuses on compiling and enhancing ETL functions to manage and analyze UEC performance data. Functions and features will continue to be added and improved over time.

## Contribution

Contributions are welcome! If you'd like to contribute, please create a pull request or open an issue to discuss improvements or changes.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
