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
        logging.info(f"Successfully loaded Excel file (first attempt): {excel_file_path}")

        df.columns = [str(col).strip() for col in df.columns]
        
        logging.debug(f"DEBUG: Columns loaded from Excel (after stripping): {df.columns.tolist()}")
        
        # Using 'Invoice #' as per user feedback
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
        logging.debug(f"Trying to load status from SQLite database: {db_path}, table: {table_name}")
        conn = sqlite3.connect(db_path)
        # Check if table exists
        query_table_exists = f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
        cursor = conn.cursor()
        cursor.execute(query_table_exists)
        table_exists = cursor.fetchone()

        if not table_exists:
            logging.info(f"Table '{table_name}' not found in database '{db_path}'. Creating a new empty DataFrame.")
            conn.close()
            empty_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
            for col in date_columns: # Ensure date columns are datetime type
                if col in empty_df.columns:
                    empty_df[col] = pd.to_datetime(empty_df[col])
            return empty_df

        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()
        logging.info(f"Successfully loaded status data from SQLite: {db_path}, table: {table_name}")

        # Convert date columns to datetime objects after loading from SQL
        for col_name in date_columns:
            if col_name in df.columns:
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
            else:
                logging.warning(f"Date column '{col_name}' not found in data loaded from SQLite table '{table_name}'.")

        # Ensure loaded data has all expected columns, adding missing ones
        for col in config.EXPECTED_COLUMNS:
            if col not in df.columns:
                logging.warning(f"Column '{col}' missing in loaded status data from SQLite. Adding it.")
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

        return df
        
    except sqlite3.Error as e:
        logging.error(f"SQLite error loading status from {db_path}, table {table_name}: {e}. Creating empty DataFrame.", exc_info=True)
        messagebox.showerror("Database Error", f"Error loading status data from database: {e}. A new empty dataset will be used.")
    except Exception as e:
        logging.error(f"Unexpected error loading status from {db_path}: {e}. Creating empty DataFrame.", exc_info=True)
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
        logging.warning("Attempted to save a None DataFrame. Operation skipped.")
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

        conn = sqlite3.connect(db_path)
        # Save DataFrame to SQL, replacing table if it exists
        df_to_save.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        
        logging.info(f"Status data successfully saved to SQLite: {db_path}, table: {table_name}")
        messagebox.showinfo("Info", "Status saved successfully to database.")
    except sqlite3.Error as e:
        logging.error(f"SQLite error saving status to {db_path}, table {table_name}: {e}", exc_info=True)
        messagebox.showerror("Database Error", f"Error saving status to database: {e}")
    except Exception as e:
        logging.error(f"Error saving status to {db_path}: {e}", exc_info=True)
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
    logging.info("Starting data processing: merging new Excel data with current status.")

    new_df_sanitized = new_df_raw.copy()
    if not all(isinstance(col, str) and col == col.strip() for col in new_df_sanitized.columns):
        new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}")

    # Using 'Invoice #' as the key column
    key_column = 'Invoice #'

    if key_column not in new_df_sanitized.columns:
        logging.error(f"Process_data: '{key_column}' column is missing in the new data. Cannot proceed with merge.")
        messagebox.showerror("Processing Error", 
                             f"The '{key_column}' column could not be found in the loaded Excel data. "
                             "Please ensure the column exists and is correctly named.")
        return None

    new_df_sanitized[key_column] = new_df_sanitized[key_column].astype(str)
    if key_column in current_status_df.columns: # Ensure current_status_df also has key_column as string
        current_status_df[key_column] = current_status_df[key_column].astype(str)


    if '#' in new_df_sanitized.columns and '#' not in config.EXPECTED_COLUMNS:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.")

    if '#' in current_status_df.columns and '#' not in config.EXPECTED_COLUMNS:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.")

    cols_from_new = [col for col in config.EXPECTED_COLUMNS 
                     if col in new_df_sanitized.columns and col not in ['Status', 'Notes']]
    
    if key_column not in cols_from_new and key_column in new_df_sanitized.columns : 
        cols_from_new.insert(0, key_column)
    elif key_column not in cols_from_new and key_column not in new_df_sanitized.columns:
        logging.error(f"Process_data: '{key_column}' is critically missing from new_df_sanitized for merge key preparation.")
        return None


    merged_df = pd.merge(
        current_status_df,
        new_df_sanitized[cols_from_new], 
        on=key_column, # Using 'Invoice #'
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}")

    processed_rows = []
    date_cols_config = ['Order Date', 'Turn in Date'] 

    for index, row in merged_df.iterrows():
        current_row_data = {}
        invoice_num = row.get(key_column) # Using 'Invoice #'

        if row['_merge'] == 'right_only':
            logging.debug(f"Processing new job (right_only): {key_column} {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                elif col == 'Status':
                    current_row_data[col] = 'New'
                elif col == 'Notes':
                    current_row_data[col] = ''
                else:
                    value = row.get(col) 
                    current_row_data[col] = value if pd.notna(value) else (pd.NaT if col in date_cols_config else None)


        elif row['_merge'] == 'left_only':
            logging.debug(f"Processing job missing from new Excel (left_only): {key_column} {invoice_num}")
            original_status_val = None
            existing_notes = ""
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                else:
                    old_col_name = col + '_old' 
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
            logging.info(f"{key_column} {invoice_num}: Status set to '{config.REVIEW_MISSING_STATUS}' and Notes updated.")

        elif row['_merge'] == 'both':
            logging.debug(f"Processing existing job (both): {key_column} {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == key_column:
                    current_row_data[col] = invoice_num
                elif col in ['Status', 'Notes']: 
                    current_row_data[col] = row.get(col) if pd.notna(row.get(col)) else ('' if col == 'Notes' else 'New')
                else: 
                    new_val = row.get(col + '_new')
                    old_val = row.get(col + '_old')
                    
                    if col in date_cols_config:
                        current_row_data[col] = new_val if pd.notna(new_val) else old_val 
                    elif pd.notna(new_val):
                        current_row_data[col] = new_val
                    else: 
                        current_row_data[col] = old_val


        for col_check in config.EXPECTED_COLUMNS:
            if col_check not in current_row_data:
                logging.warning(f"Safeguard: Column '{col_check}' was missing for {key_column} {invoice_num}. Setting default.")
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
        elif col_final_cast == key_column : # Using 'Invoice #'
             final_df[col_final_cast] = final_df[col_final_cast].astype(str)
    
    logging.info("Data processing finished.")
    return final_df