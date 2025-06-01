# data_utils.py
import pandas as pd
from tkinter import messagebox
import logging
import datetime # For timestamping the alert note
import sqlite3 # <<< ADDED for SQLite operations

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
        if isinstance(date_series, pd.Series) and not date_series.empty :
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
        logging.debug(f"LOAD_EXCEL: Raw columns from Excel (header=0 attempt): {df.columns.tolist()}") # <<< NEW DEBUG LOG
        # logging.info(f"Successfully loaded Excel file (first attempt): {excel_file_path}") # Original info log

        df.columns = [str(col).strip() for col in df.columns]
        
        logging.debug(f"LOAD_EXCEL: Stripped columns from Excel (header=0 attempt): {df.columns.tolist()}") # <<< REVISED DEBUG LOG (was DEBUG: Columns loaded from Excel (after stripping))
        
        # Using 'Invoice #' as per user feedback
        if 'Invoice #' not in df.columns:
            logging.warning(f"LOAD_EXCEL: Initial load missing 'Invoice #'. Trying header=1.") # <<< NEW DEBUG LOG
            # logging.warning(f"Initial load of {excel_file_path} missing 'Invoice #' column. Assuming an extra header row and trying again (header=1).") # Original warning
            df = pd.read_excel(excel_file_path, header=1)
            logging.debug(f"LOAD_EXCEL: Raw columns from Excel (header=1 attempt): {df.columns.tolist()}") # <<< NEW DEBUG LOG
            df.columns = [str(col).strip() for col in df.columns]
            logging.debug(f"LOAD_EXCEL: Stripped columns from Excel (header=1 attempt): {df.columns.tolist()}") # <<< REVISED DEBUG LOG (was DEBUG: Columns after attempting header=1 (after stripping))
            
            if 'Invoice #' not in df.columns:
                 logging.error(f"LOAD_EXCEL: Failed to find 'Invoice #' column even after skipping the first row in {excel_file_path}.") # <<< REVISED DEBUG LOG
                 # logging.error(f"Failed to find 'Invoice #' column even after skipping the first row in {excel_file_path}.") # Original error
                 messagebox.showerror("Excel Load Error", 
                                      f"Could not find the required 'Invoice #' column in the Excel file:\n{excel_file_path}\n\n"
                                      "Please ensure the Excel sheet has a header row containing 'Invoice #' and other expected columns, "
                                      "and that it's located in either the first or second row.")
                 return None
            else:
                logging.info(f"LOAD_EXCEL: Successfully re-loaded Excel file {excel_file_path} with header=1.") # <<< REVISED DEBUG LOG
                # logging.info(f"Successfully re-loaded Excel file {excel_file_path} with header=1.") # Original info
        else:
             logging.info(f"LOAD_EXCEL: Initial load of {excel_file_path} with header=0 looks OK (found 'Invoice #').") # <<< REVISED DEBUG LOG
             # logging.info(f"Initial load of {excel_file_path} looks OK (found 'Invoice #').") # Original info

        date_columns_to_adjust = ['Order Date', 'Turn in Date']
        current_timestamp = pd.Timestamp.now()

        for col_name in date_columns_to_adjust:
            if col_name in df.columns:
                logging.debug(f"LOAD_EXCEL: Processing Excel column '{col_name}' for date adjustment.") # <<< REVISED DEBUG LOG (was DEBUG: Processing Excel column)
                # logging.debug(f"DEBUG: Processing Excel column '{col_name}' for date adjustment.") # Original debug
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
                df[col_name] = _adjust_ambiguous_date_years(df[col_name], current_timestamp, series_name=col_name)
            else:
                logging.debug(f"LOAD_EXCEL: Date column '{col_name}' NOT found in Excel columns for adjustment.") # <<< REVISED DEBUG LOG (was DEBUG: Excel Column)
                # logging.debug(f"DEBUG: Excel Column '{col_name}' NOT found in Excel columns. Skipping adjustment for this column.") # Original debug
        
        logging.debug(f"LOAD_EXCEL: DataFrame head after loading and date adjustments:\n{df.head().to_string()}") # <<< NEW DEBUG LOG
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
    Loads the current job status DataFrame from an SQLite database (config.STATUS_FILE).
    Applies date year adjustment to specified date columns after loading (though less critical if DB stores full dates).
    Ensures that the loaded DataFrame conforms to config.EXPECTED_COLUMNS.

    Returns:
        pd.DataFrame: The loaded (or newly created and processed) status DataFrame.
    """
    db_path = config.STATUS_FILE
    table_name = config.DB_TABLE_NAME
    date_columns = ['Order Date', 'Turn in Date'] # From config or defined logic

    try:
        logging.debug(f"LOAD_STATUS: Trying to load status from SQLite database: {db_path}, table: {table_name}") # <<< REVISED DEBUG LOG
        # logging.debug(f"Trying to load status from SQLite database: {db_path}, table: {table_name}") # Original debug
        conn = sqlite3.connect(db_path)
        # Check if table exists
        query_table_exists = f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
        cursor = conn.cursor()
        cursor.execute(query_table_exists)
        table_exists = cursor.fetchone()

        if not table_exists:
            logging.info(f"LOAD_STATUS: Table '{table_name}' not found in database '{db_path}'. Creating a new empty DataFrame.") # <<< REVISED DEBUG LOG
            # logging.info(f"Table '{table_name}' not found in database '{db_path}'. Creating a new empty DataFrame.") # Original info
            conn.close()
            empty_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
            for col in date_columns: # Ensure date columns are datetime type
                if col in empty_df.columns:
                    empty_df[col] = pd.to_datetime(empty_df[col])
            return empty_df

        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()
        logging.info(f"LOAD_STATUS: Successfully loaded status data from SQLite: {db_path}, table: {table_name}. Columns: {df.columns.tolist()}") # <<< REVISED INFO LOG + NEW DEBUG
        # logging.info(f"Successfully loaded status data from SQLite: {db_path}, table: {table_name}") # Original info

        # Convert date columns to datetime objects after loading from SQL
        for col_name in date_columns:
            if col_name in df.columns:
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
            else:
                logging.warning(f"LOAD_STATUS: Date column '{col_name}' not found in data loaded from SQLite table '{table_name}'.") # <<< REVISED WARNING LOG
                # logging.warning(f"Date column '{col_name}' not found in data loaded from SQLite table '{table_name}'.") # Original warning

        # Ensure loaded data has all expected columns, adding missing ones
        for col in config.EXPECTED_COLUMNS:
            if col not in df.columns:
                logging.warning(f"LOAD_STATUS: Column '{col}' missing in loaded status data from SQLite. Adding it.") # <<< REVISED WARNING LOG
                # logging.warning(f"Column '{col}' missing in loaded status data from SQLite. Adding it.") # Original warning
                if col in date_columns:
                    df[col] = pd.NaT
                else:
                    df[col] = None
        
        # Re-ensure date columns are datetime type after potential additions and reindex
        df = df.reindex(columns=config.EXPECTED_COLUMNS)
        for col_name_date in date_columns:
            if col_name_date in df.columns:
                 df[col_name_date] = pd.to_datetime(df[col_name_date], errors='coerce')
        
        # _adjust_ambiguous_date_years might be less relevant if SQLite stores full dates
        # but can be kept if there's a chance partial dates make it into the DB somehow.
        # current_timestamp = pd.Timestamp.now()
        # for col_name in date_columns:
        #     if col_name in df.columns:
        #         df[col_name] = _adjust_ambiguous_date_years(df[col_name], current_timestamp, series_name=f"SQLite_{col_name}")
        logging.debug(f"LOAD_STATUS: DataFrame head after loading and processing SQLite data:\n{df.head().to_string()}") # <<< NEW DEBUG LOG
        return df
        
    except sqlite3.Error as e:
        logging.error(f"LOAD_STATUS: SQLite error loading status from {db_path}, table {table_name}: {e}. Creating empty DataFrame.", exc_info=True) # <<< REVISED ERROR LOG
        # logging.error(f"SQLite error loading status from {db_path}, table {table_name}: {e}. Creating empty DataFrame.", exc_info=True) # Original error
        messagebox.showerror("Database Error", f"Error loading status data from database: {e}. A new empty dataset will be used.")
    except Exception as e:
        logging.error(f"LOAD_STATUS: Unexpected error loading status from {db_path}: {e}. Creating empty DataFrame.", exc_info=True) # <<< REVISED ERROR LOG
        # logging.error(f"Unexpected error loading status from {db_path}: {e}. Creating empty DataFrame.", exc_info=True) # Original error
        messagebox.showerror("Error", f"Unexpected error loading status data: {e}. A new empty dataset will be used.")

    # Fallback: return empty DataFrame if any error occurs
    empty_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    for col in date_columns:
        if col in empty_df.columns:
            empty_df[col] = pd.to_datetime(empty_df[col])
    return empty_df


def save_status(df: pd.DataFrame) -> None:
    """
    Saves the current status DataFrame to an SQLite database (config.STATUS_FILE).
    The table (config.DB_TABLE_NAME) is replaced if it exists.
    Only the config.EXPECTED_COLUMNS are saved.

    Args:
        df (pd.DataFrame): The DataFrame containing the current job statuses to save.
    """
    if df is None:
        logging.warning("SAVE_STATUS: Attempted to save a None DataFrame. Operation skipped.") # <<< REVISED WARNING LOG
        # logging.warning("Attempted to save a None DataFrame. Operation skipped.") # Original warning
        messagebox.showwarning("Save Warning", "No data to save.")
        return

    db_path = config.STATUS_FILE
    table_name = config.DB_TABLE_NAME
    date_columns_to_check = ['Order Date', 'Turn in Date']


    try:
        df_to_save = df.copy()
        # Ensure date columns are datetime objects for SQLite compatibility (though SQLite stores them as text/real/integer)
        # Pandas to_sql handles type conversion appropriately for common types.
        for col in date_columns_to_check:
            if col in df_to_save.columns:
                df_to_save[col] = pd.to_datetime(df_to_save[col], errors='coerce')
        
        # Ensure only expected columns are saved, in the correct order.
        df_to_save = df_to_save.reindex(columns=config.EXPECTED_COLUMNS)
        logging.debug(f"SAVE_STATUS: DataFrame head before saving to SQLite:\n{df_to_save.head().to_string()}") # <<< NEW DEBUG LOG

        conn = sqlite3.connect(db_path)
        # Save DataFrame to SQL, replacing table if it exists
        df_to_save.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        
        logging.info(f"SAVE_STATUS: Status data successfully saved to SQLite: {db_path}, table: {table_name}") # <<< REVISED INFO LOG
        # logging.info(f"Status data successfully saved to SQLite: {db_path}, table: {table_name}") # Original info
        messagebox.showinfo("Info", "Status saved successfully to database.")
    except sqlite3.Error as e:
        logging.error(f"SAVE_STATUS: SQLite error saving status to {db_path}, table {table_name}: {e}", exc_info=True) # <<< REVISED ERROR LOG
        # logging.error(f"SQLite error saving status to {db_path}, table {table_name}: {e}", exc_info=True) # Original error
        messagebox.showerror("Database Error", f"Error saving status to database: {e}")
    except Exception as e:
        logging.error(f"SAVE_STATUS: Error saving status to {db_path}: {e}", exc_info=True) # <<< REVISED ERROR LOG
        # logging.error(f"Error saving status to {db_path}: {e}", exc_info=True) # Original error
        messagebox.showerror("Error", f"Error saving status: {e}")


def process_data(new_df_raw: pd.DataFrame, current_status_df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Merges new Excel data with the current status DataFrame, handling new, 
    existing, and missing jobs. Assumes new_df_raw has already had its dates adjusted by load_excel.
    Uses 'Invoice #' as the key column.

    Args:
        new_df_raw (pd.DataFrame): Raw DataFrame loaded from Excel (dates should be adjusted).
        current_status_df (pd.DataFrame): Current job status DataFrame (dates should be adjusted by load_status).

    Returns:
        pd.DataFrame | None: The processed and merged DataFrame, or None if a critical
                             error (like missing 'Invoice #' after loading) occurs.
    """
    logging.info("PROCESS_DATA: Starting data processing: merging new Excel data with current status.") # <<< REVISED INFO LOG
    logging.debug(f"PROCESS_DATA: Initial new_df_raw columns: {new_df_raw.columns.tolist()}") # <<< NEW DEBUG LOG
    logging.debug(f"PROCESS_DATA: Initial current_status_df columns: {current_status_df.columns.tolist()}") # <<< NEW DEBUG LOG
    logging.debug(f"PROCESS_DATA: current_status_df head (first 5 rows):\n{current_status_df.head().to_string()}") # <<< NEW DEBUG LOG
    # logging.info("Starting data processing: merging new Excel data with current status.") # Original info

    new_df_sanitized = new_df_raw.copy()
    original_new_columns = list(new_df_sanitized.columns) # For logging # <<< NEW DEBUG LOG
    if not all(isinstance(col, str) and col == col.strip() for col in new_df_sanitized.columns): # Check if sanitization is needed
        new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    if original_new_columns != list(new_df_sanitized.columns): # <<< NEW DEBUG LOG
        logging.debug(f"PROCESS_DATA: Sanitized new DataFrame columns from: {original_new_columns} to: {list(new_df_sanitized.columns)}") # <<< NEW DEBUG LOG
    else: # <<< NEW DEBUG LOG
        logging.debug(f"PROCESS_DATA: new_df_sanitized columns (no change after sanitization): {list(new_df_sanitized.columns)}") # <<< NEW DEBUG LOG
    # logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}") # Original debug

    # Using 'Invoice #' as the key column
    key_column = 'Invoice #'

    if key_column not in new_df_sanitized.columns:
        logging.error(f"PROCESS_DATA: '{key_column}' column is missing in the new data (new_df_sanitized). Cannot proceed with merge.") # <<< REVISED ERROR LOG
        # logging.error(f"Process_data: '{key_column}' column is missing in the new data. Cannot proceed with merge.") # Original error
        messagebox.showerror("Processing Error", 
                             f"The '{key_column}' column could not be found in the loaded Excel data. "
                             "Please ensure the column exists and is correctly named.")
        return None

    new_df_sanitized[key_column] = new_df_sanitized[key_column].astype(str)
    if key_column in current_status_df.columns: # Ensure current_status_df also has key_column as string
        current_status_df[key_column] = current_status_df[key_column].astype(str)
    else: # If current_status_df is empty or somehow missing key_column # <<< NEW DEBUG LOG
        logging.warning(f"PROCESS_DATA: Key column '{key_column}' not found in current_status_df. If it's initially empty, this is fine.") # <<< NEW DEBUG LOG
        # Ensure it exists for the merge if current_status_df is empty # <<< NEW DEBUG LOG
        if current_status_df.empty and key_column not in current_status_df.columns: # <<< NEW DEBUG LOG
             current_status_df[key_column]=pd.Series(dtype='object') # <<< NEW DEBUG LOG


    if '#' in new_df_sanitized.columns and '#' not in config.EXPECTED_COLUMNS:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("PROCESS_DATA: Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.") # <<< REVISED DEBUG LOG
        # logging.debug("Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.") # Original debug

    if '#' in current_status_df.columns and '#' not in config.EXPECTED_COLUMNS:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("PROCESS_DATA: Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.") # <<< REVISED DEBUG LOG
        # logging.debug("Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.") # Original debug

    cols_from_new = [] # <<< NEW DEBUG LOG
    for col_name in config.EXPECTED_COLUMNS: # <<< NEW DEBUG LOG
        if col_name in new_df_sanitized.columns: # <<< NEW DEBUG LOG
            if col_name not in ['Status', 'Notes']: # <<< NEW DEBUG LOG
                cols_from_new.append(col_name) # <<< NEW DEBUG LOG
                logging.debug(f"PROCESS_DATA: Added '{col_name}' to cols_from_new.") # <<< NEW DEBUG LOG
        else: # <<< NEW DEBUG LOG
            logging.debug(f"PROCESS_DATA: Column '{col_name}' from EXPECTED_COLUMNS not found in new_df_sanitized.columns, not added to cols_from_new.") # <<< NEW DEBUG LOG
    # cols_from_new = [col for col in config.EXPECTED_COLUMNS 
    #                  if col in new_df_sanitized.columns and col not in ['Status', 'Notes']] # Original list comprehension
    
    # Ensure key_column is in cols_from_new if it's in new_df_sanitized # <<< NEW DEBUG LOG
    if key_column in new_df_sanitized.columns and key_column not in cols_from_new: # <<< REVISED/NEW DEBUG LOG
    # if key_column not in cols_from_new and key_column in new_df_sanitized.columns : # Original if
        cols_from_new.insert(0, key_column)
        logging.debug(f"PROCESS_DATA: Ensured '{key_column}' is in cols_from_new.") # <<< NEW DEBUG LOG
    elif key_column not in cols_from_new and key_column not in new_df_sanitized.columns: # Should have been caught # <<< NEW DEBUG LOG
        logging.error(f"PROCESS_DATA: '{key_column}' is critically missing from new_df_sanitized for merge key preparation (cols_from_new).") # <<< REVISED ERROR LOG
        # logging.error(f"Process_data: '{key_column}' is critically missing from new_df_sanitized for merge key preparation.") # Original error
        return None
    
    # Remove duplicates from cols_from_new just in case, though logic should prevent it # <<< NEW DEBUG LOG
    cols_from_new = sorted(list(set(cols_from_new)), key=cols_from_new.index) # <<< NEW DEBUG LOG
    logging.debug(f"PROCESS_DATA: Final cols_from_new for merge: {cols_from_new}") # <<< NEW DEBUG LOG
    
    if not cols_from_new: # <<< NEW DEBUG LOG
        logging.error("PROCESS_DATA: cols_from_new is empty. This likely means no matching columns from Excel for merging (excluding Status/Notes). Check Excel headers and config.EXPECTED_COLUMNS.") # <<< NEW DEBUG LOG
        # For now, let it proceed, merge might just become effectively a copy of current_status_df + new empty records # <<< NEW DEBUG LOG
    
    valid_cols_for_merge = [col for col in cols_from_new if col in new_df_sanitized.columns] # <<< NEW DEBUG LOG
    if not valid_cols_for_merge: # <<< NEW DEBUG LOG
        logging.warning("PROCESS_DATA: No valid columns from 'cols_from_new' are actually in 'new_df_sanitized'. Merging with effectively empty right side.") # <<< NEW DEBUG LOG
        # Create an empty df with just the key column if needed by merge logic # <<< NEW DEBUG LOG
        new_data_for_merge = pd.DataFrame(columns=[key_column]) if key_column not in new_df_sanitized.columns else new_df_sanitized[[key_column]].copy() # <<< NEW DEBUG LOG
    else: # <<< NEW DEBUG LOG
        new_data_for_merge = new_df_sanitized[valid_cols_for_merge] # <<< NEW DEBUG LOG

    logging.debug(f"PROCESS_DATA: Columns in new_data_for_merge (right side of merge): {new_data_for_merge.columns.tolist()}") # <<< NEW DEBUG LOG


    merged_df = pd.merge(
        current_status_df,
        # new_df_sanitized[cols_from_new], # Original selection
        new_data_for_merge, # <<< Use the carefully prepared new_data_for_merge
        on=key_column, # Using 'Invoice #'
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"PROCESS_DATA: Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}") # <<< REVISED DEBUG LOG
    logging.debug(f"PROCESS_DATA: merged_df columns: {merged_df.columns.tolist()}") # <<< NEW DEBUG LOG
    logging.debug(f"PROCESS_DATA: merged_df head (first 5 rows):\n{merged_df.head().to_string()}") # <<< NEW DEBUG LOG
    # logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}") # Original debug

    processed_rows = []
    date_cols_config = ['Order Date', 'Turn in Date'] 

    for index, row_series in merged_df.iterrows(): # Renamed 'row' to 'row_series' for clarity
        current_row_data = {}
        invoice_num = row_series.get(key_column) # Using 'Invoice #'
        merge_type = row_series.get('_merge') # <<< Get merge type
        logging.debug(f"PROCESS_DATA: Processing Invoice: {invoice_num}, MergeType: {merge_type}") # <<< NEW DEBUG LOG

        if merge_type == 'right_only': # <<< Use merge_type variable
            logging.debug(f"  RIGHT_ONLY Branch for Invoice: {invoice_num}") # <<< NEW DEBUG LOG
            # logging.debug(f"Processing new job (right_only): {key_column} {invoice_num}") # Original debug
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                elif col == 'Status':
                    current_row_data[col] = 'New'
                elif col == 'Notes':
                    current_row_data[col] = ''
                else:
                    # CRITICAL FIX for blank columns in new invoices:
                    value_from_new_df = row_series.get(col + '_new') # <<< Get from the new data side
                    is_not_na = pd.notna(value_from_new_df) # <<< NEW DEBUG LOG
                    logging.debug(f"    RIGHT_ONLY Col: {col}, Suffix: _new, RawValue: '{value_from_new_df}', IsNotNA: {is_not_na}") # <<< NEW DEBUG LOG
                    current_row_data[col] = value_from_new_df if is_not_na else \
                                            (pd.NaT if col in date_cols_config else None)
            logging.debug(f"  RIGHT_ONLY generated current_row_data for {invoice_num}: {current_row_data}") # <<< NEW DEBUG LOG


        elif merge_type == 'left_only': # <<< Use merge_type variable
            logging.debug(f"  LEFT_ONLY Branch for Invoice: {invoice_num}") # <<< NEW DEBUG LOG
            # logging.debug(f"Processing job missing from new Excel (left_only): {key_column} {invoice_num}") # Original debug
            original_status_val = None
            existing_notes = ""
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                else:
                    old_col_name = col + '_old' 
                    val_from_row = row_series.get(old_col_name) if old_col_name in row_series else row_series.get(col) # Use row_series
                    logging.debug(f"    LEFT_ONLY Col: {col}, Value from row_series: '{val_from_row}'") # <<< NEW DEBUG LOG
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
            logging.info(f"  LEFT_ONLY: Invoice {invoice_num} Status set to '{config.REVIEW_MISSING_STATUS}'.") # <<< REVISED INFO LOG (was f"{key_column} {invoice_num}: Status set to...")
            # logging.info(f"{key_column} {invoice_num}: Status set to '{config.REVIEW_MISSING_STATUS}' and Notes updated.") # Original info
            logging.debug(f"  LEFT_ONLY generated current_row_data for {invoice_num}: {current_row_data}") # <<< NEW DEBUG LOG

        elif merge_type == 'both': # <<< Use merge_type variable
            logging.debug(f"  BOTH Branch for Invoice: {invoice_num}") # <<< NEW DEBUG LOG
            # logging.debug(f"Processing existing job (both): {key_column} {invoice_num}") # Original debug
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                elif col in ['Status', 'Notes']: 
                    val_status_notes = row_series.get(col) # Use row_series
                    logging.debug(f"    BOTH Col(Status/Notes): {col}, Value: '{val_status_notes}' (from current data)") # <<< NEW DEBUG LOG
                    current_row_data[col] = val_status_notes if pd.notna(val_status_notes) else ('' if col == 'Notes' else 'New') # Use val_status_notes
                else: 
                    new_val = row_series.get(col + '_new') # Use row_series
                    old_val = row_series.get(col + '_old') # Use row_series
                    is_new_val_not_na = pd.notna(new_val) # <<< NEW DEBUG LOG
                    logging.debug(f"    BOTH Col: {col}, NewVal: '{new_val}', OldVal: '{old_val}', IsNewNotNA: {is_new_val_not_na}") # <<< NEW DEBUG LOG
                    
                    if col in date_cols_config:
                        current_row_data[col] = new_val if is_new_val_not_na else old_val # Use is_new_val_not_na
                    elif is_new_val_not_na: # Use is_new_val_not_na
                        current_row_data[col] = new_val
                    else: 
                        current_row_data[col] = old_val
            logging.debug(f"  BOTH generated current_row_data for {invoice_num}: {current_row_data}") # <<< NEW DEBUG LOG


        # Safeguard check (already present, good)
        for col_check in config.EXPECTED_COLUMNS:
            if col_check not in current_row_data:
                logging.warning(f"PROCESS_DATA: Safeguard! Column '{col_check}' was missing for Invoice {invoice_num}. Setting default.") # <<< REVISED WARNING LOG
                # logging.warning(f"Safeguard: Column '{col_check}' was missing for {key_column} {invoice_num}. Setting default.") # Original warning
                default_value = 'New' if col_check == 'Status' else ('' if col_check == 'Notes' else None)
                if col_check in date_cols_config:
                     default_value = pd.NaT
                current_row_data[col_check] = default_value
        
        processed_rows.append(current_row_data)

    if not processed_rows:
        logging.info("PROCESS_DATA: No rows to process after merge. Returning empty DataFrame with expected columns.") # <<< REVISED INFO LOG
        # logging.info("No rows to process after merge. Returning empty DataFrame with expected columns.") # Original info
        final_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)
        logging.info(f"PROCESS_DATA: Successfully processed {len(final_df)} rows.") # <<< REVISED INFO LOG
        # logging.info(f"Successfully processed {len(final_df)} rows.") # Original info

    final_df = final_df.reindex(columns=config.EXPECTED_COLUMNS)
    for col_final_cast in config.EXPECTED_COLUMNS:
        if col_final_cast in date_cols_config:
            final_df[col_final_cast] = pd.to_datetime(final_df[col_final_cast], errors='coerce')
        elif col_final_cast == key_column : # Using 'Invoice #'
             final_df[col_final_cast] = final_df[col_final_cast].astype(str)
    
    logging.debug(f"PROCESS_DATA: Final DataFrame head before returning (first 5 rows):\n{final_df.head().to_string()}") # <<< NEW DEBUG LOG
    logging.info("PROCESS_DATA: Data processing finished.") # <<< REVISED INFO LOG
    # logging.info("Data processing finished.") # Original info
    return final_df