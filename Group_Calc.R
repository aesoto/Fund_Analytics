# Version: August 2025
# YTD metrics need to be raw, not annualized


library('quantmod')
library('PerformanceAnalytics')
library('openxlsx')
library('dplyr')
library('tidyr')
library('data.table')
library('writexl')
library('readr')
library('lubridate')


original.directory <- getwd()

# Functions for capture
up_capture_ratio <- function(Ra, Rb) {
  up_periods <- which(Rb > 0)
  apply(Ra[up_periods, ], 2, mean, na.rm = TRUE) / mean(Rb[up_periods], na.rm = TRUE)
}

down_capture_ratio <- function(Ra, Rb) {
  down_periods <- which(Rb < 0)
  apply(Ra[down_periods, ], 2, mean, na.rm = TRUE) / mean(Rb[down_periods], na.rm = TRUE)
}

# Function to clean and process sheet data
process_sheet_data <- function(filename, sheet_number) {
  # Read and clean the data
  data <- read.xlsx(filename, sheet = sheet_number, colNames = FALSE)
  data <- data[-c(1:4, 6:7), ]
  
  # Set column names
  new_names <- unlist(data[1, ], use.names = FALSE)
  setnames(data, new_names)
  data <- data[-1, ]
  
  # Clean column names
  names(data)[names(data) == "X1"] <- "date"
  data <- data[, !grepl("^X", names(data)) | names(data) == "date"]
  data <- data[complete.cases(data), ]
  
  # Process dates
  data$date <- sub(".* to ", "", data$date)
  data$date <- as.Date(data$date, format = "%m/%d/%Y")
  
  # Convert character columns to numeric (percentages to decimals)
  char_cols <- sapply(data, is.character)
  data[, char_cols] <- lapply(data[, char_cols], function(x) as.numeric(x) / 100)
  
  # Rename index columns
  names(data)[names(data) == "2000"] <- "R2K"  # Russell 2000
  names(data)[names(data) == "3000"] <- "R3K"  # Russell 3000 (if present)
  
  return(data)
}

# Function to add risk-free rate and create XTS object
add_risk_free_rate <- function(data, rf_rate) {
  # Join with risk-free rate data
  data_with_rf <- data %>%
    left_join(rf_rate, by = "date") %>%
    arrange(date) %>%
    fill(DGS3MO, .direction = "down") %>%
    mutate(rf = (DGS3MO / 252)) %>%  # Daily risk-free rate
    mutate(rf = replace_na(rf, 0)) %>%
    select(-DGS3MO)
  
  # Separate risk-free rate for XTS conversion
  rf_rate_subset <- data_with_rf %>% select(date, rf)
  data_clean <- data_with_rf %>% select(-rf)
  
  # Create XTS object for risk-free rate
  rf_rate_xts <- xts(
    x = rf_rate_subset$rf,
    order.by = rf_rate_subset$date
  )
  colnames(rf_rate_xts) <- "Rf"
  
  return(list(data = data_clean, rf_xts = rf_rate_xts))
}

# Get risk-free rate data (moved up to avoid repetition)
get_risk_free_rate <- function(start_date, end_date) {
  start_str <- format(start_date, "%Y-%m-%d")
  end_str <- format(end_date, "%Y-%m-%d")
  
  fred_url <- paste0("https://fred.stlouisfed.org/graph/fredgraph.csv?bgcolor=%23e1e9f0&chart_type=line&drp=0&fo=open%20sans&graph_bgcolor=%23ffffff&height=450&mode=fred&recession_bars=on&txtcolor=%23444444&ts=12&tts=12&width=1168&nt=0&thu=0&trc=0&show_legend=yes&show_axis_titles=yes&show_tooltip=yes&id=DGS3MO&scale=left&cosd=", 
                     start_str, "&coed=", end_str, 
                     "&line_color=%234572a7&link_values=false&line_style=solid&mark_type=none&mw=3&lw=2&ost=-99999&oet=99999&mma=0&fml=a&fq=Daily&fam=avg&fgst=lin&fgsnd=2020-02-01&line_index=1&transformation=lin&vintage_date=2025-08-08&revision_date=2025-08-08&nd=1981-09-01")
  
  rf_rate <- tryCatch({
    fred_data <- read_csv(fred_url, show_col_types = FALSE)
    rf_rate <- fred_data %>%
      rename(date = observation_date) %>%
      filter(!is.na(DGS3MO)) %>%
      mutate(DGS3MO = DGS3MO / 100)  # Convert % to decimal
    rf_rate
  }, error = function(e) {
    warning("Failed to download FRED data: ", e$message)
    warning("Falling back to quantmod method...")
    
    # Fallback to quantmod
    getSymbols("DGS3MO", src = "FRED", 
               from = start_date, to = end_date,
               periodicity = 'daily', auto.assign = FALSE, warnings = FALSE)
    rf_rate_xts <- na.omit(get("DGS3MO")) / 100
    rf_rate <- data.frame(date = index(rf_rate_xts), DGS3MO = coredata(rf_rate_xts))
    rf_rate
  })
  
  # Validate that we got data
  if (nrow(rf_rate) == 0) {
    stop("No risk-free rate data available for the specified date range")
  }
  
  return(rf_rate)
}

############# Analytic Functions
# Helper function to create empty reporting data frame
create_empty_dataframe <- function(data) {
  accounts <- colnames(data)
  accounts <- accounts[accounts != "date"]
  accounts <- gsub("-", "_", accounts)
  
  empty_data <- setNames(rep(list(double()), length(accounts)), accounts)
  
  result_df <- data.frame(Metric = character(), empty_data,
                          stringsAsFactors = FALSE,
                          check.names = FALSE)
  return(result_df)
}

# Function to convert data to xts format
convert_to_xts <- function(data) {
  data_xts <- xts(
    x = data[, -which(names(data) == "date")],  
    order.by = data$date
  )
  return(data_xts)
}

# Function to calculate basic performance metrics
calculate_basic_metrics <- function(data_xts, rf_rate_xts, period_type, scale_factor = 252) {
  metrics_df <- create_empty_dataframe(data.frame(data_xts))
  
  # Returns calculation based on period type
  if (period_type == "qtd") {
    returns <- Return.cumulative(data_xts, geometric = TRUE)
    metric_name <- "QTD Return"
  } else if (period_type == "ytd") {
    returns <- Return.cumulative(data_xts, geometric = TRUE)
    metric_name <- "YTD Return"
  } else if (period_type == "yr") {
    returns <- Return.cumulative(data_xts, geometric = TRUE)
    metric_name <- "1Yr Return"
  } else if (period_type == "5yr") {
    # UPDATED: Annualize 5-year returns
    returns <- Return.annualized(data_xts, scale = scale_factor, geometric = TRUE)
    metric_name <- "5Yr Return (Annualized)"
  } else if (period_type == "10yr") {
    # UPDATED: Annualize 10-year returns
    returns <- Return.annualized(data_xts, scale = scale_factor, geometric = TRUE)
    metric_name <- "10Yr Return (Annualized)"
  }
  
  # Format returns
  returns_df <- as.data.frame(returns)
  rownames(returns_df) <- NULL
  colnames(returns_df) <- gsub("-", "_", colnames(returns_df))
  returns_df <- returns_df %>%
    mutate(Metric = metric_name) %>%
    select(Metric, everything())
  metrics_df <- rbind(metrics_df, returns_df)
  
  # Standard Deviation calculation based on period type
  if (period_type %in% c("qtd", "ytd", "yr")) {
    # Use regular StdDev for shorter periods
    std_dev <- StdDev(data_xts)
  } else if (period_type %in% c("5yr", "10yr")) {
    # UPDATED: Annualize standard deviation for multi-year periods
    std_dev <- StdDev.annualized(data_xts, scale = scale_factor)
  } else {
    std_dev <- StdDev.annualized(data_xts, scale = scale_factor)
  }
  
  std_dev_df <- as.data.frame(std_dev)
  rownames(std_dev_df) <- NULL
  colnames(std_dev_df) <- gsub("-", "_", colnames(std_dev_df))
  std_dev_df <- std_dev_df %>%
    mutate(Metric = "Std Deviation") %>%
    select(Metric, everything())
  metrics_df <- rbind(metrics_df, std_dev_df)
  
  # Sharpe Ratio calculation based on period type
  if (period_type %in% c("qtd", "ytd", "yr")) {
    # Use non-annualized Sharpe for shorter periods
    sharpe <- SharpeRatio(data_xts, Rf = rf_rate_xts, FUN = "StdDev", annualize = FALSE)
  } else if (period_type %in% c("5yr", "10yr")) {
    # UPDATED: Annualize Sharpe ratio for multi-year periods
    sharpe <- SharpeRatio.annualized(data_xts, Rf = rf_rate_xts, scale = scale_factor, geometric = TRUE)
  } else {
    sharpe <- SharpeRatio.annualized(as.data.frame(data_xts), Rf = rf_rate_xts, 
                                     scale = scale_factor, geometric = TRUE)
  }
  
  sharpe_df <- as.data.frame(sharpe)
  rownames(sharpe_df) <- NULL
  colnames(sharpe_df) <- gsub("-", "_", colnames(sharpe_df))
  sharpe_df <- sharpe_df %>%
    mutate(Metric = "Sharpe") %>%
    select(Metric, everything())
  metrics_df <- rbind(metrics_df, sharpe_df)
  
  return(metrics_df)
}

# Function to get benchmark mapping for funds
get_benchmark_mapping <- function(data_cols) {
  # Define which funds use which benchmarks - FIXED to match actual column names
  fund_benchmark_map <- list(
    "GROWTH" = "SP50",      # Changed from "SP500" to "SP50"
    "BALANCED" = "SP50",    # Changed from "SP500" to "SP50"
    "SMALLCAP" = "R2K"
  )
  
  # Get available funds and benchmarks in the data
  available_funds <- intersect(names(fund_benchmark_map), data_cols)
  
  # Create mapping of fund indices to benchmark indices
  benchmark_mapping <- list()
  for (fund in available_funds) {
    fund_idx <- which(data_cols == fund)
    benchmark_name <- fund_benchmark_map[[fund]]
    benchmark_idx <- which(data_cols == benchmark_name)
    
    if (length(fund_idx) > 0 && length(benchmark_idx) > 0) {
      benchmark_mapping[[fund]] <- list(
        fund_idx = fund_idx,
        benchmark_idx = benchmark_idx,
        benchmark_name = benchmark_name
      )
    }
  }
  
  return(benchmark_mapping)
}

# Function to calculate advanced analytics with multiple benchmarks
calculate_advanced_metrics <- function(data_xts, rf_rate_data, period_type, scale_factor = 252) {
  # Get column names (excluding date)
  data_cols <- colnames(data_xts)
  
  # Get benchmark mapping
  benchmark_mapping <- get_benchmark_mapping(data_cols)
  
  if (length(benchmark_mapping) == 0) {
    warning("No valid fund-benchmark pairs found")
    return(list())
  }
  
  # Initialize single rows for each metric - one row per metric type
  n_cols <- length(data_cols)
  
  # Create single rows for each metric
  act_premium_row <- rep(NA, n_cols)
  trck_error_row <- rep(NA, n_cols)
  inf_ratio_row <- rep(NA, n_cols)
  correlation_row <- rep(NA, n_cols)
  alpha_row <- rep(NA, n_cols)
  beta_row <- rep(NA, n_cols)
  up_capture_row <- rep(NA, n_cols)
  down_capture_row <- rep(NA, n_cols)
  
  # Calculate metrics for each fund-benchmark pair
  for (fund_name in names(benchmark_mapping)) {
    mapping <- benchmark_mapping[[fund_name]]
    fund_idx <- mapping$fund_idx
    benchmark_idx <- mapping$benchmark_idx
    benchmark_name <- mapping$benchmark_name
    
    # Active Premium
    if (period_type %in% c("qtd", "ytd", "yr")) {
      # Use cumulative returns for shorter periods
      returns <- Return.cumulative(data_xts, geometric = TRUE)
      act_premium_val <- returns[, fund_idx] - returns[, benchmark_idx]
    } else if (period_type %in% c("5yr", "10yr")) {
      # UPDATED: Use annualized active premium for multi-year periods
      act_premium_val <- ActivePremium(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    } else {
      act_premium_val <- ActivePremium(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    }
    act_premium_row[fund_idx] <- as.numeric(act_premium_val)
    
    # Tracking Error
    if (period_type %in% c("qtd", "ytd", "yr")) {
      # Use period tracking error for shorter periods
      trck_error_val <- TrackingError(data_xts[, fund_idx], data_xts[, benchmark_idx])
    } else if (period_type %in% c("5yr", "10yr")) {
      # UPDATED: Use annualized tracking error for multi-year periods
      trck_error_val <- TrackingError(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    } else {
      trck_error_val <- TrackingError(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    }
    trck_error_row[fund_idx] <- as.numeric(trck_error_val)
    
    # Information Ratio
    if (period_type %in% c("qtd", "ytd", "yr")) {
      # Use period information ratio for shorter periods
      inf_ratio_val <- InformationRatio(data_xts[, fund_idx], data_xts[, benchmark_idx])
    } else if (period_type %in% c("5yr", "10yr")) {
      # UPDATED: Use annualized information ratio for multi-year periods
      inf_ratio_val <- InformationRatio(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    } else {
      inf_ratio_val <- InformationRatio(data_xts[, fund_idx], data_xts[, benchmark_idx], scale = scale_factor)
    }
    inf_ratio_row[fund_idx] <- as.numeric(inf_ratio_val)
    
    # Correlation - NOT annualized (correlation is dimensionless)
    correlation_val <- cor(data_xts[, fund_idx], data_xts[, benchmark_idx], 
                           use = "pairwise.complete.obs", method = "pearson")
    correlation_row[fund_idx] <- as.numeric(correlation_val)
    
    # CAPM Alpha
    if (period_type %in% c("qtd", "ytd", "yr")) {
      # Use period alpha for shorter periods
      camp_alpha_val <- CAPM.alpha(Ra = data_xts[, fund_idx],
                                   Rb = data_xts[, benchmark_idx],
                                   Rf = rf_rate_data)
    } else if (period_type %in% c("5yr", "10yr")) {
      # UPDATED: Use annualized alpha for multi-year periods
      camp_alpha_val <- CAPM.alpha(Ra = data_xts[, fund_idx],    # <- FIXED: camp -> camp
                                   Rb = data_xts[, benchmark_idx],
                                   Rf = rf_rate_data)
      # Scale alpha to annualized
      camp_alpha_val <- camp_alpha_val * scale_factor              # <- FIXED: camp -> camp
    }
    alpha_row[fund_idx] <- as.numeric(camp_alpha_val)             # <- FIXED: camp -> camp
    
    # CAPM Beta - NOT annualized (beta is dimensionless)
    camp_beta_val <- CAPM.beta(Ra = data_xts[, fund_idx],
                               Rb = data_xts[, benchmark_idx],
                               Rf = rf_rate_data)
    beta_row[fund_idx] <- as.numeric(camp_beta_val)
    
    # Up Capture - NOT annualized (capture ratios are dimensionless)
    up_capture_val <- up_capture_ratio(data_xts[, fund_idx, drop = FALSE], data_xts[, benchmark_idx])
    up_capture_row[fund_idx] <- as.numeric(up_capture_val)
    
    # Down Capture - NOT annualized (capture ratios are dimensionless)
    down_capture_val <- down_capture_ratio(data_xts[, fund_idx, drop = FALSE], data_xts[, benchmark_idx])
    down_capture_row[fund_idx] <- as.numeric(down_capture_val)
  }
  
  # Create matrices with single rows and proper names
  results_list <- list(
    ActivePremium = matrix(act_premium_row, nrow = 1, dimnames = list("Active Premium", data_cols)),
    TrackingError = matrix(trck_error_row, nrow = 1, dimnames = list("Tracking Error", data_cols)),
    InformationRatio = matrix(inf_ratio_row, nrow = 1, dimnames = list("Information Ratio", data_cols)),
    Correlation = matrix(correlation_row, nrow = 1, dimnames = list("Correlation", data_cols)),
    Alpha = matrix(alpha_row, nrow = 1, dimnames = list("Alpha", data_cols)),
    Beta = matrix(beta_row, nrow = 1, dimnames = list("Beta", data_cols)),
    upCapture = matrix(up_capture_row, nrow = 1, dimnames = list("Up Capture", data_cols)),
    downCapture = matrix(down_capture_row, nrow = 1, dimnames = list("Down Capture", data_cols))
  )
  
  return(results_list)
}

# Function to format advanced metrics into data frame
format_advanced_metrics <- function(results_list, template_df) {
  if (length(results_list) == 0) {
    return(data.frame())
  }
  
  # Get the column structure from template
  template_cols <- names(template_df)
  
  # Initialize final dataframe with proper structure
  final_df <- data.frame(Metric = character(0), stringsAsFactors = FALSE)
  
  # Add all columns from template except Metric
  for (col_name in template_cols[-1]) {  # Skip first column (Metric)
    final_df[[col_name]] <- numeric(0)
  }
  
  # Process each metric type and add as a single row
  for (metric_name in names(results_list)) {
    metric_matrix <- results_list[[metric_name]]
    
    # Convert matrix to data frame row
    metric_row <- as.data.frame(metric_matrix, stringsAsFactors = FALSE)
    metric_row$Metric <- rownames(metric_matrix)[1]  # Get the metric name from row names
    
    # Reorder columns to match template (Metric first)
    metric_row <- metric_row[, c("Metric", setdiff(names(metric_row), "Metric"))]
    
    # Ensure column names match template structure
    names(metric_row) <- gsub("-", "_", names(metric_row))
    
    # Add missing columns if needed and set to appropriate values
    for (col_name in template_cols) {
      if (!col_name %in% names(metric_row)) {
        if (col_name %in% c("R3k", "R3K")) {
          metric_row[[col_name]] <- "--"
        } else {
          metric_row[[col_name]] <- NA
        }
      }
    }
    
    # Reorder to match template column order
    metric_row <- metric_row[, template_cols]
    
    # Bind to final result
    final_df <- rbind(final_df, metric_row)
  }
  
  # Reset row names
  rownames(final_df) <- NULL
  
  return(final_df)
}

# Function to calculate prorated fees
calculate_prorated_fees <- function(end_date, period_type, data_columns) {
  # Create base fee values matching the data structure
  # Fund columns: GROWTH, BALANCED, SMALLCAP + benchmark columns: SP50, R2K
  fee_values <- c("Est. Fee")  # Start with metric name
  
  # Add fees for each column in the data
  for (col_name in data_columns) {
    if (col_name == "date") next  # Skip date column
    
    # Set fees based on column type - Updated with correct fee values
    if (col_name %in% c("GROWTH", "BALANCED", "SMALLCAP")) {
      # Fund fees with correct annual rates
      if (col_name == "GROWTH") fee_values <- c(fee_values, 0.0062)    # 0.62%
      else if (col_name == "BALANCED") fee_values <- c(fee_values, 0.0071)  # 0.71%
      else if (col_name == "SMALLCAP") fee_values <- c(fee_values, 0.0094)  # 0.94%
    } else {
      # Benchmark columns (SP50, R2K) get "--"
      fee_values <- c(fee_values, "--")
    }
  }
  
  # Calculate period fraction and metric name based on period type
  if (period_type == "qtd") {
    # For QTD, use 1/4 of annual fee
    period_fraction <- 0.25
    metric_name <- "Est. Fee (QTD)"
  } else if (period_type == "ytd") {
    # Calculate year fraction for YTD proration
    start_of_year <- as.Date(format(end_date, "%Y-01-01"))
    days_in_year <- as.numeric(as.Date(format(end_date, "%Y-12-31")) - start_of_year + 1)
    days_elapsed <- as.numeric(end_date - start_of_year + 1)
    period_fraction <- days_elapsed / days_in_year
    metric_name <- "Est. Fee (YTD Prorated)"
  } else if (period_type == "yr") {
    # Annual fees - use full annual fee
    period_fraction <- 1.0
    metric_name <- "Est. Fee"
  } else if (period_type %in% c("5yr", "10yr")) {
    # Multi-year periods - use annualized fee (same as annual)
    period_fraction <- 1.0
    if (period_type == "5yr") {
      metric_name <- "Est. Fee (5Yr Annualized)"
    } else {
      metric_name <- "Est. Fee (10Yr Annualized)"
    }
  } else {
    # Default case
    period_fraction <- 1.0
    metric_name <- "Est. Fee"
  }
  
  fee_values[1] <- metric_name
  return(list(values = fee_values, fraction = period_fraction))
}

# Create fee rows
create_fee_row <- function(fee_data, template_df, period_fraction) {
  new_row <- setNames(as.data.frame(t(fee_data$values), stringsAsFactors = FALSE), names(template_df))
  
  # Apply period fraction to all fund fees (not just when < 1)
  fee_cols <- setdiff(names(new_row), c("Metric"))
  
  # Only process numeric fee values, skip "--" values
  for (col in fee_cols) {
    if (col %in% names(new_row) && !is.na(suppressWarnings(as.numeric(new_row[[col]])))) {
      # Apply the period fraction
      prorated_fee <- as.numeric(new_row[[col]]) * period_fraction
      
      # Format appropriately - use regular decimal format instead of scientific notation
      new_row[[col]] <- prorated_fee
    }
  }
  
  return(new_row)
}

######### Formatting Functions
# Function to create Excel styles
create_excel_styles <- function() {
  list(
    title_style = createStyle(fontSize = 14, textDecoration = "bold"),
    date_style = createStyle(fontSize = 12, textDecoration = "bold"),
    header_style = createStyle(halign = "center", textDecoration = "bold"),
    char_style = createStyle(halign = "left"),
    percent_style = createStyle(numFmt = "0.00%"),
    decimal_style = createStyle(numFmt = "0.00"),
    right_align_style = createStyle(halign = "right"),
    fee_style = createStyle(numFmt = "0.000%")  # For fees with 3 decimal places
  )
}

# Function to determine row formatting rules based on period
get_formatting_rules <- function(period_type) {
  base_rules <- list(
    percent_rows = c(1, 2, 4, 5, 8, 10, 11, 12),  # Returns, Active Premium, Correlation, Capture ratios
    decimal_rows = c(3, 6, 7, 9),                 # Std Dev, Tracking Error, Info Ratio, Alpha, Beta
    fee_row = NULL
  )
  
  # Adjust for different periods (fee row position may vary)
  if (period_type %in% c("qtd", "ytd", "yr", "5yr", "10yr")) {
    base_rules$fee_row <- 13  # Assuming 12 metrics + 1 fee row
  }
  
  return(base_rules)
}

# Function to prepare data for Excel export
prepare_data_for_excel <- function(df) {
  # Safety check: ensure all columns have the same number of rows
  col_lengths <- sapply(df, length)
  if (any(col_lengths != nrow(df))) {
    cat("WARNING: Column length mismatch detected!\n")
    for (i in seq_along(col_lengths)) {
      if (col_lengths[i] != nrow(df)) {
        col_name <- names(col_lengths)[i]
        cat("Column", col_name, "has length", col_lengths[i], "but dataframe has", nrow(df), "rows\n")
        # Fix the column by filling with appropriate values
        if (col_name %in% c("R3k", "R3K")) {
          df[[col_name]] <- rep("--", nrow(df))
        } else {
          df[[col_name]] <- rep(NA, nrow(df))
        }
      }
    }
  }
  
  # Convert appropriate columns to numeric
  cols_to_numeric <- setdiff(names(df), c("Metric", "R3k"))
  cols_to_numeric <- cols_to_numeric[cols_to_numeric %in% names(df)]
  
  if (length(cols_to_numeric) > 0) {
    df[cols_to_numeric] <- lapply(df[cols_to_numeric], function(x) {
      # Handle scientific notation and convert to numeric
      if (is.character(x)) {
        numeric_vals <- suppressWarnings(as.numeric(x))
        return(numeric_vals)
      } else {
        return(as.numeric(x))
      }
    })
  }
  
  # Handle R3k column properly - only if it exists
  if ("R3k" %in% names(df)) {
    # Convert R3k column, preserving "--" values as NA
    df$R3k <- ifelse(df$R3k == "--", NA, suppressWarnings(as.numeric(df$R3k)))
  }
  
  return(df)
}

# Function to add and format sheet
add_formatted_sheet <- function(wb, df, sheet_name, report_date, period_type = "ytd") {
  # Prepare data
  df <- prepare_data_for_excel(df)
  
  # Get styles and formatting rules
  styles <- create_excel_styles()
  formatting_rules <- get_formatting_rules(period_type)
  
  # Add worksheet
  addWorksheet(wb, sheet_name)
  
  # Add title and date
  writeData(wb, sheet_name, x = "Taft Hartley Accounts Performance", startCol = 1, startRow = 1)
  addStyle(wb, sheet_name, style = styles$title_style, rows = 1, cols = 1, stack = TRUE)
  
  formatted_date <- format(report_date, "%d %B %Y")
  period_label <- switch(period_type,
                         "qtd" = "Quarter-to-Date",
                         "ytd" = "Year-to-Date", 
                         "yr" = "1-Year",
                         "5yr" = "5-Year",
                         "10yr" = "10-Year",
                         "Period")
  
  writeData(wb, sheet_name, x = paste("As of:", formatted_date, "-", period_label), 
            startCol = 1, startRow = 2)
  addStyle(wb, sheet_name, style = styles$date_style, rows = 2, cols = 1, stack = TRUE)
  
  # Write the data with headers
  writeData(wb, sheet_name, df, startCol = 1, startRow = 4, colNames = TRUE, 
            headerStyle = styles$header_style)
  
  # Format column widths
  setColWidths(wb, sheet_name, cols = 1, widths = 25) # Metric column (wider for longer names)
  
  # Get column positions
  all_col_positions <- seq_along(df)
  excluded_positions <- which(names(df) %in% c("Metric", "R3k"))
  data_cols <- setdiff(all_col_positions, excluded_positions)
  r3k_col <- which(tolower(names(df)) == "r3k")
  
  if (length(data_cols) > 0) {
    setColWidths(wb, sheet_name, cols = data_cols, widths = 12)
  }
  
  # Handle NA values in R3k column
  na_rows <- which(is.na(df$R3k))
  if (length(na_rows) > 0 && length(r3k_col) > 0) {
    excel_na_rows <- na_rows + 4
    addStyle(wb, sheet_name, style = styles$right_align_style, 
             rows = excel_na_rows, cols = r3k_col, stack = TRUE)
  }
  
  # Apply formatting based on row content
  for (i in seq_len(nrow(df))) {
    row_index_excel <- i + 4
    
    if (i %in% formatting_rules$percent_rows) {
      # Percentage formatting for returns, active premium, correlations, capture ratios
      addStyle(wb, sheet_name, style = styles$percent_style,
               rows = row_index_excel, cols = c(data_cols, r3k_col), stack = TRUE)
    } else if (i %in% formatting_rules$decimal_rows) {
      # Decimal formatting for std dev, tracking error, info ratio, alpha, beta
      addStyle(wb, sheet_name, style = styles$decimal_style,
               rows = row_index_excel, cols = c(data_cols, r3k_col), stack = TRUE)
    } else if (!is.null(formatting_rules$fee_row) && i == formatting_rules$fee_row) {
      # Special formatting for fees (3 decimal places)
      addStyle(wb, sheet_name, style = styles$fee_style,
               rows = row_index_excel, cols = data_cols, stack = TRUE)
    }
  }
  
  # Left-align Metric column
  addStyle(wb, sheet_name, style = styles$char_style,
           rows = 5:(nrow(df)+4), cols = 1, gridExpand = TRUE, stack = TRUE)
}

# Function to generate output filename and path
generate_output_path <- function(filename, base_dir = "C:/Users/asoto/mairsandpower.com/Quant - Performance Analytics/Funds") {
  date_str <- sub(".*_(\\d{8})\\.xlsx", "\\1", filename)
  report_date <- as.Date(date_str, format = "%Y%m%d")
  
  output_filename <- paste0("Funds_Report_", format(report_date, "%Y%m%d"), ".xlsx")
  output_path <- file.path(base_dir, output_filename)
  
  return(list(path = output_path, date = report_date))
}

# Function to create complete workbook
create_performance_workbook <- function(accounts_list, filename, output_dir = NULL) {
  # Generate output path and report date
  if (is.null(output_dir)) {
    output_dir <- "C:/Users/asoto/mairsandpower.com/Quant - Performance Analytics/Funds"
  }
  
  output_info <- generate_output_path(filename, output_dir)
  
  # Create workbook
  wb <- createWorkbook()
  
  # Add summary sheet first
  cat("Adding summary dashboard...\n")
  add_summary_sheet(wb, accounts_list, output_info$date)
  
  # Add sheets for each period (these will be positioned after summary)
  if ("qtd" %in% names(accounts_list)) {
    cat("Adding QTD sheet...\n")
    add_formatted_sheet(wb, accounts_list$qtd, "QTD", output_info$date, "qtd")
  }
  
  if ("ytd" %in% names(accounts_list)) {
    cat("Adding YTD sheet...\n")
    add_formatted_sheet(wb, accounts_list$ytd, "YTD", output_info$date, "ytd")
  }
  
  if ("yr" %in% names(accounts_list)) {
    cat("Adding 1Yr sheet...\n")
    add_formatted_sheet(wb, accounts_list$yr, "1Yr", output_info$date, "yr")
  }
  
  if ("5yr" %in% names(accounts_list)) {
    cat("Adding 5Yr sheet...\n")
    add_formatted_sheet(wb, accounts_list$"5yr", "5Yr", output_info$date, "5yr")
  }
  
  if ("10yr" %in% names(accounts_list)) {
    cat("Adding 10Yr sheet...\n")
    add_formatted_sheet(wb, accounts_list$"10yr", "10Yr", output_info$date, "10yr")
  }
  
  return(list(workbook = wb, path = output_info$path, date = output_info$date))
}

# Function to add enhanced summary dashboard
add_summary_sheet <- function(wb, accounts_list, report_date) {
  addWorksheet(wb, "Summary")  # Add summary sheet
  
  styles <- create_excel_styles()
  
  # Add title
  writeData(wb, "Summary", x = "Performance Summary Dashboard", startCol = 1, startRow = 1)
  addStyle(wb, "Summary", style = styles$title_style, rows = 1, cols = 1, stack = TRUE)
  
  formatted_date <- format(report_date, "%d %B %Y")
  writeData(wb, "Summary", x = paste("As of:", formatted_date), startCol = 1, startRow = 2)
  addStyle(wb, "Summary", style = styles$date_style, rows = 2, cols = 1, stack = TRUE)
  
  # Create simplified summary table
  start_row <- 4
  
  # Get ALL column names (funds AND benchmarks)
  sample_df <- accounts_list[[1]]  # Use first available dataset
  all_column_names <- setdiff(names(sample_df), "Metric")
  
  # Include ALL columns: funds AND benchmarks (GROWTH, BALANCED, SMALLCAP, SP50, R2K)
  fund_and_benchmark_names <- intersect(all_column_names, c("GROWTH", "BALANCED", "SMALLCAP", "SP50", "R2K"))
  
  # Create the summary data frame with all return periods
  summary_data <- data.frame(
    Return = c("QTD Return", "YTD Return", "1Yr Return", "5Yr Return", "10Yr Return"),
    stringsAsFactors = FALSE
  )
  
  # Add columns for each fund AND benchmark
  for (col_name in fund_and_benchmark_names) {
    summary_data[[col_name]] <- NA
  }
  
  # Fill in the data from each period - FIXED return labels to match actual data
  periods <- c("qtd", "ytd", "yr", "5yr", "10yr")
  return_labels <- c("QTD Return", "YTD Return", "1Yr Return", "5Yr Return (Annualized)", "10Yr Return (Annualized)")
  
  for (i in seq_along(periods)) {
    period <- periods[i]
    return_label <- return_labels[i]
    
    if (period %in% names(accounts_list)) {
      df <- accounts_list[[period]]
      returns_row <- df[df$Metric == return_label, ]
      
      if (nrow(returns_row) > 0) {
        # Extract values for each fund AND benchmark
        for (col_name in fund_and_benchmark_names) {
          if (col_name %in% names(returns_row)) {
            summary_data[i, col_name] <- returns_row[[col_name]]
          }
        }
      }
    }
  }
  
  # Prepare data for Excel (convert fund AND benchmark columns to numeric)
  data_cols <- setdiff(names(summary_data), "Return")
  summary_data[data_cols] <- lapply(summary_data[data_cols], function(x) {
    if (is.character(x)) {
      # Handle "--" values and convert to numeric
      numeric_vals <- suppressWarnings(as.numeric(x))
      return(numeric_vals)
    } else {
      as.numeric(x)
    }
  })
  
  # Write the data to Excel
  writeData(wb, "Summary", summary_data, startCol = 1, startRow = start_row, 
            colNames = TRUE, headerStyle = styles$header_style)
  
  # Format the table
  # Set column widths
  setColWidths(wb, "Summary", cols = 1, widths = 15)  # Return column
  if (length(data_cols) > 0) {
    setColWidths(wb, "Summary", cols = 2:(length(data_cols) + 1), widths = 12)  # All data columns
  }
  
  # Apply percentage formatting to all data cells (funds AND benchmarks)
  data_rows <- (start_row + 1):(start_row + nrow(summary_data))
  data_col_positions <- which(names(summary_data) %in% data_cols)
  
  if (length(data_col_positions) > 0) {
    addStyle(wb, "Summary", style = styles$percent_style,
             rows = data_rows, cols = data_col_positions, gridExpand = TRUE, stack = TRUE)
  }
  
  # Left-align the Return column
  addStyle(wb, "Summary", style = styles$char_style,
           rows = data_rows, cols = 1, gridExpand = TRUE, stack = TRUE)
}


############# End of Functions
# Main execution
# Get data
# ToDo: pull latest return document from revised location
setwd("C:/Users/asoto/mairsandpower.com/Quant - Performance Analytics/Funds/Return_Data")
data.files <- list.files(pattern = 'Funds_Rets')
data.files <- data.files %>% sort
filename <- tail(data.files, 1)

# Validate file exists and show which file is being used
if (!file.exists(filename)) {
  stop("Data file not found: ", filename)
}
cat("Using data file:", filename, "\n")

# Process all five sheets
cat("Processing data sheets...\n")
data.qtd <- process_sheet_data(filename, 1)
data.ytd <- process_sheet_data(filename, 2)
data.yr <- process_sheet_data(filename, 3)
data.5yr <- process_sheet_data(filename, 4)
data.10yr <- process_sheet_data(filename, 5)

# Validate that all sheets have data
if (any(sapply(list(data.qtd, data.ytd, data.yr, data.5yr, data.10yr), nrow) == 0)) {
  stop("One or more data sheets are empty")
}

# Get date range for risk-free rate (using the widest range)
all_dates <- c(data.qtd$date, data.ytd$date, data.yr$date, data.5yr$date, data.10yr$date)
startDate <- min(all_dates, na.rm = TRUE)
endDate <- max(all_dates, na.rm = TRUE)

# Get risk-free rate data once
cat("Downloading risk-free rate data from", format(startDate, "%Y-%m-%d"), "to", format(endDate, "%Y-%m-%d"), "...\n")
rf_rate <- get_risk_free_rate(startDate, endDate)

# Add risk-free rate to each dataset
cat("Adding risk-free rates to datasets...\n")
qtd_result <- add_risk_free_rate(data.qtd, rf_rate)
ytd_result <- add_risk_free_rate(data.ytd, rf_rate)
yr_result <- add_risk_free_rate(data.yr, rf_rate)
fiveyear_result <- add_risk_free_rate(data.5yr, rf_rate)
tenyear_result <- add_risk_free_rate(data.10yr, rf_rate)

# Extract results
data.qtd <- qtd_result$data
data.ytd <- ytd_result$data
data.yr <- yr_result$data
data.5yr <- fiveyear_result$data
data.10yr <- tenyear_result$data

rf_rate_xts.qtd <- qtd_result$rf_xts
rf_rate_xts.ytd <- ytd_result$rf_xts
rf_rate_xts.yr <- yr_result$rf_xts
rf_rate_xts.5yr <- fiveyear_result$rf_xts
rf_rate_xts.10yr <- tenyear_result$rf_xts

cat("Data processing complete!\n")
cat("QTD data: ", nrow(data.qtd), " rows, ", ncol(data.qtd), " columns\n")
cat("YTD data: ", nrow(data.ytd), " rows, ", ncol(data.ytd), " columns\n")
cat("1Yr data: ", nrow(data.yr), " rows, ", ncol(data.yr), " columns\n")
cat("5Yr data: ", nrow(data.5yr), " rows, ", ncol(data.5yr), " columns\n")
cat("10Yr data: ", nrow(data.10yr), " rows, ", ncol(data.10yr), " columns\n")

cat("Setting up reporting data frames...\n")

# Create empty data frames
accounts.qtd <- create_empty_dataframe(data.qtd)
accounts.ytd <- create_empty_dataframe(data.ytd)
accounts.yr <- create_empty_dataframe(data.yr)
accounts.5yr <- create_empty_dataframe(data.5yr)
accounts.10yr <- create_empty_dataframe(data.10yr)

# Transform data to xts objects
cat("Converting data to XTS format...\n")
data_xts.qtd <- convert_to_xts(data.qtd)
data_xts.ytd <- convert_to_xts(data.ytd)
data_xts.yr <- convert_to_xts(data.yr)
data_xts.5yr <- convert_to_xts(data.5yr)
data_xts.10yr <- convert_to_xts(data.10yr)

# Data quality check
cat("Data quality check:\n")
cat("QTD date range:", format(min(index(data_xts.qtd)), "%Y-%m-%d"), "to", format(max(index(data_xts.qtd)), "%Y-%m-%d"), "\n")
cat("YTD date range:", format(min(index(data_xts.ytd)), "%Y-%m-%d"), "to", format(max(index(data_xts.ytd)), "%Y-%m-%d"), "\n")
cat("1Yr date range:", format(min(index(data_xts.yr)), "%Y-%m-%d"), "to", format(max(index(data_xts.yr)), "%Y-%m-%d"), "\n")
cat("5Yr date range:", format(min(index(data_xts.5yr)), "%Y-%m-%d"), "to", format(max(index(data_xts.5yr)), "%Y-%m-%d"), "\n")
cat("10Yr date range:", format(min(index(data_xts.10yr)), "%Y-%m-%d"), "to", format(max(index(data_xts.10yr)), "%Y-%m-%d"), "\n")

# Calculate basic metrics for all periods
cat("Calculating basic performance metrics...\n")
accounts.qtd <- rbind(accounts.qtd, calculate_basic_metrics(data_xts.qtd, rf_rate_xts.qtd, "qtd"))
accounts.ytd <- rbind(accounts.ytd, calculate_basic_metrics(data_xts.ytd, rf_rate_xts.ytd, "ytd"))
accounts.yr <- rbind(accounts.yr, calculate_basic_metrics(data_xts.yr, rf_rate_xts.yr, "yr"))
accounts.5yr <- rbind(accounts.5yr, calculate_basic_metrics(data_xts.5yr, rf_rate_xts.5yr, "5yr"))
accounts.10yr <- rbind(accounts.10yr, calculate_basic_metrics(data_xts.10yr, rf_rate_xts.10yr, "10yr"))

# Calculate advanced metrics for all periods
cat("Calculating advanced analytics...\n")
advanced_qtd <- calculate_advanced_metrics(data_xts.qtd, rf_rate_xts.qtd, "qtd")
advanced_ytd <- calculate_advanced_metrics(data_xts.ytd, rf_rate_xts.ytd, "ytd")
advanced_yr <- calculate_advanced_metrics(data_xts.yr, rf_rate_xts.yr, "yr")
advanced_5yr <- calculate_advanced_metrics(data_xts.5yr, rf_rate_xts.5yr, "5yr")
advanced_10yr <- calculate_advanced_metrics(data_xts.10yr, rf_rate_xts.10yr, "10yr")

# Format and append advanced metrics
final_df.qtd <- format_advanced_metrics(advanced_qtd, accounts.qtd)
final_df.ytd <- format_advanced_metrics(advanced_ytd, accounts.ytd)
final_df.yr <- format_advanced_metrics(advanced_yr, accounts.yr)
final_df.5yr <- format_advanced_metrics(advanced_5yr, accounts.5yr)
final_df.10yr <- format_advanced_metrics(advanced_10yr, accounts.10yr)

accounts.qtd <- rbind(accounts.qtd, final_df.qtd)
accounts.ytd <- rbind(accounts.ytd, final_df.ytd)
accounts.yr <- rbind(accounts.yr, final_df.yr)
accounts.5yr <- rbind(accounts.5yr, final_df.5yr)
accounts.10yr <- rbind(accounts.10yr, final_df.10yr)

# Calculate and append fees
cat("Calculating fees...\n")
qtd_fees <- calculate_prorated_fees(endDate, "qtd", names(data.qtd))
ytd_fees <- calculate_prorated_fees(endDate, "ytd", names(data.ytd))
yr_fees <- calculate_prorated_fees(endDate, "yr", names(data.yr))
fiveyear_fees <- calculate_prorated_fees(endDate, "5yr", names(data.5yr))
tenyear_fees <- calculate_prorated_fees(endDate, "10yr", names(data.10yr))

# Append fee rows
qtd_fee_row <- create_fee_row(qtd_fees, accounts.qtd, qtd_fees$fraction)
ytd_fee_row <- create_fee_row(ytd_fees, accounts.ytd, ytd_fees$fraction)
yr_fee_row <- create_fee_row(yr_fees, accounts.yr, yr_fees$fraction)
fiveyear_fee_row <- create_fee_row(fiveyear_fees, accounts.5yr, fiveyear_fees$fraction)
tenyear_fee_row <- create_fee_row(tenyear_fees, accounts.10yr, tenyear_fees$fraction)

accounts.qtd <- rbind(accounts.qtd, qtd_fee_row)
accounts.ytd <- rbind(accounts.ytd, ytd_fee_row)
accounts.yr <- rbind(accounts.yr, yr_fee_row)
accounts.5yr <- rbind(accounts.5yr, fiveyear_fee_row)
accounts.10yr <- rbind(accounts.10yr, tenyear_fee_row)

cat("Performance analytics complete!\n")
cat("QTD metrics: ", nrow(accounts.qtd), " rows\n")
cat("YTD metrics: ", nrow(accounts.ytd), " rows\n")
cat("1Yr metrics: ", nrow(accounts.yr), " rows\n")
cat("5Yr metrics: ", nrow(accounts.5yr), " rows\n")
cat("10Yr metrics: ", nrow(accounts.10yr), " rows\n")

# Format results and save
cat("Preparing Excel export...\n")

# Organize accounts data
accounts_list <- list(
  qtd = accounts.qtd,
  ytd = accounts.ytd,
  yr = accounts.yr,
  "5yr" = accounts.5yr,
  "10yr" = accounts.10yr
)

# Create workbook
wb_info <- create_performance_workbook(accounts_list, filename)

# Save workbook
cat("Saving workbook to:", wb_info$path, "\n")
saveWorkbook(wb_info$workbook, file = wb_info$path, overwrite = TRUE)

cat("Excel report generated successfully!\n")
cat("Report date:", format(wb_info$date, "%d %B %Y"), "\n")
cat("File saved to:", wb_info$path, "\n")

cat("\n=== EXECUTION SUMMARY ===\n")
cat("Data processed for", length(accounts_list), "periods\n")
cat("Total metrics calculated:", sum(sapply(accounts_list, nrow)), "\n")
cat("File size:", file.size(wb_info$path), "bytes\n")
cat("========================\n")
