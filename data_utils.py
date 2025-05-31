# data_utils.py
import pandas as pd
from tkinter import messagebox
import logging
import datetime # For timestamping the alert note

# Import configurations from config.py
import config

# Helper function to adjust year for ambiguously parsed dates
def _adjust_ambiguous_date_years(date_series: pd.Series, current_timestamp: pd.Timestamp, series_name: str = "Unknown") -> pd.Series:
    """
    Adjusts years for dates in a Series that might have been ambiguously parsed.
    If a date was parsed into the current year but is later than the current date,
    it's assumed to be from the previous year.

    Args:
        date_series (pd.Series): A pandas Series already converted to datetime objects (with errors='coerce').
        current_timestamp (pd.Timestamp): The current timestamp to compare against.
        series_name (str): The name of the series/column, for logging purposes.

    Returns:
        pd.Series: The date Series with adjusted years where applicable.
    """
    if not isinstance(date_series, pd.Series) or date_series.empty or not pd.api.types.is_datetime64_any_dtype(date_series):
        # If not a Series, empty, or not datetime, return as is.
        # This handles cases where a column might be missing or already processed to non-datetime by an error.
        if isinstance(date_series, pd.Series) and not date_series.empty : # Log only if it was a series but not datetime
             logging.debug(f"Series '{series_name}' is not of datetime type or is empty. Skipping year adjustment.")
        return date_series

    adjusted_series = date_series.copy()
    mask = (adjusted_series.dt.year == current_timestamp.year) & \
           (adjusted_series > current_timestamp) & \
           (adjusted_series.notna())

    if mask.any():
        adjusted_series.loc[mask] = adjusted_series.loc[mask] - pd.DateOffset(years=1)
        logging.info(f"Adjusted year for some dates in series '{series_name}' assuming they were from the previous year.")
    
    return adjusted_series


def load_excel(excel_file_path: str) -> pd.DataFrame | None:
    """
    Loads data from an Excel file into a Pandas DataFrame.
    Tries to read normally, and if key columns (like 'Invoice #') are missing,
    tries again skipping the first row, assuming it might be an extra title row.
    Adjusts year for date columns if they appear to be future dates from a yearless source.

    Args:
        excel_file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame | None: A DataFrame containing the Excel data,
                             or None if an error occurs during loading.
    """
    try:
        logging.debug(f"Attempting to load Excel file from: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        logging.info(f"Successfully loaded Excel file (first attempt): {excel_file_path}")

        df.columns = [str(col).strip() for col in df.columns]
        
        logging.debug(f"DEBUG: Columns loaded from Excel (after stripping): {df.columns.tolist()}")
        
        if 'Invoice #' not in df.columns:
            logging.warning(f"Initial load of {excel_file_path} missing 'Invoice #' column. Assuming an extra header row and trying again (header=1).")
            df = pd.read_excel(excel_file_path, header=1)
            df.columns = [str(col).strip() for col in df.columns]
            logging.debug(f"DEBUG: Columns after attempting header=1 (after stripping): {df.columns.tolist()}")
            
            if 'Invoice #' not in df.columns:
                 logging.error(f"Failed to find 'Invoice #' column even after skipping the first row in {excel_file_path}.")
                 messagebox.showerror("Excel Load Error", 
                                      f"Could not find the required 'Invoice #' column in the Excel file:\n{excel_file_path}\n\n"
                                      "Please ensure the Excel sheet has a header row containing 'Invoice #' and other expected columns, "
                                      "and that it's located in either the first or second row.")
                 return None
            else:
                logging.info(f"Successfully re-loaded Excel file {excel_file_path} with header=1.")
        else:
             logging.info(f"Initial load of {excel_file_path} looks OK (found 'Invoice #').")

        date_columns_to_adjust = ['Order Date', 'Turn in Date']
        current_timestamp = pd.Timestamp.now()

        for col_name in date_columns_to_adjust:
            if col_name in df.columns:
                logging.debug(f"DEBUG: Processing Excel column '{col_name}' for date adjustment.")
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
                df[col_name] = _adjust_ambiguous_date_years(df[col_name], current_timestamp, series_name=col_name)
            else:
                logging.debug(f"DEBUG: Excel Column '{col_name}' NOT found in Excel columns. Skipping adjustment for this column.")
        
        return df

    except FileNotFoundError:
        logging.error(f"Excel file not found at path: {excel_file_path}")
        messagebox.showerror("Error", f"File not found: {excel_file_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading Excel file {excel_file_path}: {e}", exc_info=True)
        messagebox.showerror("Error", f"An unexpected error occurred while reading the Excel file: {e}")
        return None


def load_status() -> pd.DataFrame:
    """
    Loads the current job status DataFrame from a pickle file (config.STATUS_FILE).
    Applies date year adjustment to specified date columns after loading.
    Ensures that the loaded DataFrame conforms to config.EXPECTED_COLUMNS.

    Returns:
        pd.DataFrame: The loaded (or newly created and processed) status DataFrame.
    """
    try:
        logging.debug(f"Trying to load status from {config.STATUS_FILE}")
        df = pd.read_pickle(config.STATUS_FILE)
        logging.info(f"Successfully loaded status data from {config.STATUS_FILE}")

        # --- Apply date year adjustment logic after loading from pickle ---
        date_columns_to_adjust = ['Order Date', 'Turn in Date']
        current_timestamp = pd.Timestamp.now()

        for col_name in date_columns_to_adjust:
            if col_name in df.columns:
                logging.debug(f"DEBUG: Processing Pickle column '{col_name}' for date adjustment.")
                # Ensure the column is datetime, as pickle might store objects or mixed types
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
                
                # Apply the year adjustment using the helper function
                df[col_name] = _adjust_ambiguous_date_years(df[col_name], current_timestamp, series_name=col_name)
            else:
                # This case should be less common if the pickle file was saved correctly by this app
                logging.warning(f"DEBUG: Date column '{col_name}' for adjustment not found in loaded pickle data. Will be created if in EXPECTED_COLUMNS.")
        # --- End: Date year adjustment logic ---

        # Ensure loaded data has all expected columns, adding missing ones with None (or NaT for dates)
        for col in config.EXPECTED_COLUMNS:
            if col not in df.columns:
                logging.warning(f"Column '{col}' missing in loaded status data. Adding it.")
                if col in date_columns_to_adjust:
                    df[col] = pd.NaT # Initialize missing date columns as NaT
                else:
                    df[col] = None # Initialize other missing columns as None
        
        # Re-ensure date columns are datetime type after potential additions
        for col_name_date in date_columns_to_adjust:
            if col_name_date in df.columns:
                 df[col_name_date] = pd.to_datetime(df[col_name_date], errors='coerce')


        # Return DataFrame with columns in the defined order and only expected columns
        # This reindex also ensures that if a column was added (e.g. missing date column), it's included
        df = df.reindex(columns=config.EXPECTED_COLUMNS) 
        return df
        
    except FileNotFoundError:
        logging.info(f"Status file '{config.STATUS_FILE}' not found. Creating a new empty DataFrame.")
        # Create an empty DataFrame with expected columns and types
        empty_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
        for col in config.EXPECTED_COLUMNS:
            if col in ['Order Date', 'Turn in Date']: # Specify your date columns
                empty_df[col] = pd.to_datetime(empty_df[col])
        return empty_df
    except Exception as e:
        logging.error(f"Error loading status from {config.STATUS_FILE}: {e}. Creating empty DataFrame.", exc_info=True)
        messagebox.showerror("Error", f"Error loading status data: {e}. A new empty dataset will be used.")
        empty_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
        for col in config.EXPECTED_COLUMNS:
            if col in ['Order Date', 'Turn in Date']: # Specify your date columns
                empty_df[col] = pd.to_datetime(empty_df[col])
        return empty_df


def save_status(df: pd.DataFrame) -> None:
    """
    Saves the current status DataFrame to a pickle file (config.STATUS_FILE).

    Only the config.EXPECTED_COLUMNS are saved, ensuring a consistent schema.

    Args:
        df (pd.DataFrame): The DataFrame containing the current job statuses to save.
    """
    if df is None:
        logging.warning("Attempted to save a None DataFrame. Operation skipped.")
        messagebox.showwarning("Save Warning", "No data to save.")
        return
    try:
        # Before saving, ensure date columns are indeed datetime objects.
        # This helps maintain type consistency in the pickle file.
        df_to_save = df.copy()
        date_cols = ['Order Date', 'Turn in Date']
        for col in date_cols:
            if col in df_to_save.columns:
                df_to_save[col] = pd.to_datetime(df_to_save[col], errors='coerce')
        
        # Ensure only expected columns are saved, in the correct order.
        df_to_save = df_to_save.reindex(columns=config.EXPECTED_COLUMNS)
        df_to_save.to_pickle(config.STATUS_FILE)
        logging.info(f"Status data successfully saved to {config.STATUS_FILE}")
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        logging.error(f"Error saving status to {config.STATUS_FILE}: {e}", exc_info=True)
        messagebox.showerror("Error", f"Error saving status: {e}")


def process_data(new_df_raw: pd.DataFrame, current_status_df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Merges new Excel data with the current status DataFrame, handling new, 
    existing, and missing jobs. Assumes new_df_raw has already had its dates adjusted by load_excel.

    Args:
        new_df_raw (pd.DataFrame): Raw DataFrame loaded from Excel (dates should be adjusted).
        current_status_df (pd.DataFrame): Current job status DataFrame (dates should be adjusted by load_status).

    Returns:
        pd.DataFrame | None: The processed and merged DataFrame, or None if a critical
                             error (like missing 'Invoice #' after loading) occurs.
    """
    logging.info("Starting data processing: merging new Excel data with current status.")

    new_df_sanitized = new_df_raw.copy()
    # Column sanitization is already done in load_excel, but if called directly, ensure it
    if not all(isinstance(col, str) and col == col.strip() for col in new_df_sanitized.columns):
        new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}")

    if 'Invoice #' not in new_df_sanitized.columns:
        logging.error("Process_data: 'Invoice #' column is missing in the new data. Cannot proceed with merge.")
        messagebox.showerror("Processing Error", 
                             "The 'Invoice #' column could not be found in the loaded Excel data. "
                             "Please ensure the column exists and is correctly named.")
        return None

    new_df_sanitized['Invoice #'] = new_df_sanitized['Invoice #'].astype(str)
    # Ensure current_status_df also has Invoice # as string, might be redundant if load_status handles it but safe
    if 'Invoice #' in current_status_df.columns:
        current_status_df['Invoice #'] = current_status_df['Invoice #'].astype(str)


    if '#' in new_df_sanitized.columns and '#' not in config.EXPECTED_COLUMNS:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.")

    if '#' in current_status_df.columns and '#' not in config.EXPECTED_COLUMNS:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.")

    cols_from_new = [col for col in config.EXPECTED_COLUMNS 
                     if col in new_df_sanitized.columns and col not in ['Status', 'Notes']]
    
    if 'Invoice #' not in cols_from_new and 'Invoice #' in new_df_sanitized.columns : # Ensure Invoice # is in new_df_sanitized
        cols_from_new.insert(0, 'Invoice #')
    elif 'Invoice #' not in cols_from_new and 'Invoice #' not in new_df_sanitized.columns:
        logging.error("Process_data: 'Invoice #' is critically missing from new_df_sanitized for merge key preparation.")
        return None


    merged_df = pd.merge(
        current_status_df,
        new_df_sanitized[cols_from_new], 
        on='Invoice #',
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}")

    processed_rows = []
    date_cols_config = ['Order Date', 'Turn in Date'] # From config or defined

    for index, row in merged_df.iterrows():
        current_row_data = {}
        invoice_num = row.get('Invoice #')

        if row['_merge'] == 'right_only':
            logging.debug(f"Processing new job (right_only): Invoice # {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                elif col == 'Status':
                    current_row_data[col] = 'New'
                elif col == 'Notes':
                    current_row_data[col] = ''
                else:
                    value = row.get(col) # This 'col' directly refers to columns from new_df_sanitized part of merge
                    current_row_data[col] = value if pd.notna(value) else (pd.NaT if col in date_cols_config else None)


        elif row['_merge'] == 'left_only':
            logging.debug(f"Processing job missing from new Excel (left_only): Invoice # {invoice_num}")
            original_status_val = None
            existing_notes = ""
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                else:
                    old_col_name = col + '_old' # This suffix applies if col was in cols_from_new
                    # If col was NOT in cols_from_new (like Status, Notes), it won't have _old suffix from merge
                    # It will have its original name from current_status_df
                    val_from_row = row.get(old_col_name) if old_col_name in row else row.get(col)
                    current_row_data[col] = val_from_row if pd.notna(val_from_row) else (pd.NaT if col in date_cols_config else None)

                    if col == 'Status': original_status_val = current_row_data[col]
                    if col == 'Notes': existing_notes = str(current_row_data[col]) if current_row_data[col] is not None else ""
            
            current_row_data['Status'] = config.REVIEW_MISSING_STATUS
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            alert_message = f"System Alert ({timestamp}): Job not in last Excel import."
            if original_status_val and original_status_val != config.REVIEW_MISSING_STATUS:
                alert_message += f" Previous status: '{original_status_val}'. "
            alert_message += "Verify if closed (set Status to 'Closed') or if active (re-include in Excel & update status)."
            current_row_data['Notes'] = (alert_message + "\n-----\n" + existing_notes).strip()
            logging.info(f"Invoice # {invoice_num}: Status set to '{config.REVIEW_MISSING_STATUS}' and Notes updated.")

        elif row['_merge'] == 'both':
            logging.debug(f"Processing existing job (both): Invoice # {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                elif col in ['Status', 'Notes']: # These come from the 'left' side (current_status_df)
                    current_row_data[col] = row.get(col) if pd.notna(row.get(col)) else ('' if col == 'Notes' else 'New')
                else: # These columns were part of cols_from_new, so they have _new and _old suffixes
                    new_val = row.get(col + '_new')
                    old_val = row.get(col + '_old')
                    
                    if col in date_cols_config:
                        current_row_data[col] = new_val if pd.notna(new_val) else old_val # pd.notna handles NaT correctly
                    elif pd.notna(new_val):
                        current_row_data[col] = new_val
                    else: # new_val is None or NaN (for non-datetime)
                        current_row_data[col] = old_val


        for col_check in config.EXPECTED_COLUMNS:
            if col_check not in current_row_data:
                logging.warning(f"Safeguard: Column '{col_check}' was missing for Invoice # {invoice_num}. Setting default.")
                default_value = 'New' if col_check == 'Status' else ('' if col_check == 'Notes' else None)
                if col_check in date_cols_config:
                     default_value = pd.NaT
                current_row_data[col_check] = default_value
        
        processed_rows.append(current_row_data)

    if not processed_rows:
        logging.info("No rows to process after merge. Returning empty DataFrame with expected columns.")
        final_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)
        logging.info(f"Successfully processed {len(final_df)} rows.")

    final_df = final_df.reindex(columns=config.EXPECTED_COLUMNS)
    for col_final_cast in config.EXPECTED_COLUMNS:
        if col_final_cast in date_cols_config:
            final_df[col_final_cast] = pd.to_datetime(final_df[col_final_cast], errors='coerce')
        elif col_final_cast == 'Invoice #' :
             final_df[col_final_cast] = final_df[col_final_cast].astype(str)
    
    logging.info("Data processing finished.")
    return final_df