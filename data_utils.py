# data_utils.py
import pandas as pd
from tkinter import messagebox
import logging
import datetime # For timestamping the alert note

# Import configurations from config.py
import config

def load_excel(excel_file_path: str) -> pd.DataFrame | None:
    """
    Loads data from an Excel file into a Pandas DataFrame.
    Tries to read normally, and if key columns (like 'Invoice #') are missing,
    tries again skipping the first row, assuming it might be an extra title row.

    Args:
        excel_file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame | None: A DataFrame containing the Excel data,
                             or None if an error occurs during loading.
    """
    try:
        logging.debug(f"Attempting to load Excel file from: {excel_file_path}")
        # --- First Attempt: Read with default header (row 0) ---
        df = pd.read_excel(excel_file_path)
        logging.info(f"Successfully loaded Excel file (first attempt): {excel_file_path}")

        # Clean and check column names
        df.columns = [str(col).strip() for col in df.columns]
        
        # Check if 'Invoice #' exists. It's a critical column for processing.
        if 'Invoice #' not in df.columns:
            logging.warning(f"Initial load of {excel_file_path} missing 'Invoice #' column. Assuming an extra header row and trying again (header=1).")
            # --- Second Attempt: Read skipping the first row (header=1) ---
            df = pd.read_excel(excel_file_path, header=1)
            df.columns = [str(col).strip() for col in df.columns] # Clean columns again
            
            # Check 'Invoice #' again after the second attempt
            if 'Invoice #' not in df.columns:
                 logging.error(f"Failed to find 'Invoice #' column even after skipping the first row in {excel_file_path}.")
                 messagebox.showerror("Excel Load Error", 
                                      f"Could not find the required 'Invoice #' column in the Excel file:\n{excel_file_path}\n\n"
                                      "Please ensure the Excel sheet has a header row containing 'Invoice #' and other expected columns, "
                                      "and that it's located in either the first or second row.")
                 return None # Return None if still not found
            else:
                logging.info(f"Successfully re-loaded Excel file {excel_file_path} with header=1.")
        else:
             logging.info(f"Initial load of {excel_file_path} looks OK (found 'Invoice #').")

        return df

    except FileNotFoundError:
        logging.error(f"Excel file not found at path: {excel_file_path}")
        messagebox.showerror("Error", f"File not found: {excel_file_path}")
        return None
    except Exception as e:
        # Catch other potential errors during reading (e.g., corrupted file, no sheet)
        logging.error(f"Error reading Excel file {excel_file_path}: {e}", exc_info=True)
        messagebox.showerror("Error", f"An unexpected error occurred while reading the Excel file: {e}")
        return None


def load_status() -> pd.DataFrame:
    """
    Loads the current job status DataFrame from a pickle file (config.STATUS_FILE).

    If the file is not found, or if there's an error loading it,
    an empty DataFrame with the config.EXPECTED_COLUMNS structure is returned.
    Ensures that the loaded DataFrame conforms to config.EXPECTED_COLUMNS,
    adding any missing columns with None values and ensuring correct order.

    Returns:
        pd.DataFrame: The loaded (or newly created) status DataFrame.
    """
    try:
        logging.debug(f"Trying to load status from {config.STATUS_FILE}")
        df = pd.read_pickle(config.STATUS_FILE)
        logging.info(f"Successfully loaded status data from {config.STATUS_FILE}")

        # Ensure loaded data has all expected columns, adding missing ones with None
        for col in config.EXPECTED_COLUMNS:
            if col not in df.columns:
                logging.warning(f"Column '{col}' missing in loaded status data. Adding it with None values.")
                df[col] = None
        # Return DataFrame with columns in the defined order and only expected columns
        return df.reindex(columns=config.EXPECTED_COLUMNS) # Ensures order and drops extra columns
    except FileNotFoundError:
        logging.info(f"Status file '{config.STATUS_FILE}' not found. Creating a new empty DataFrame.")
        return pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    except Exception as e:
        logging.error(f"Error loading status from {config.STATUS_FILE}: {e}. Creating empty DataFrame.", exc_info=True)
        messagebox.showerror("Error", f"Error loading status data: {e}. A new empty dataset will be used.")
        return pd.DataFrame(columns=config.EXPECTED_COLUMNS)


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
        # Ensure only expected columns are saved, in the correct order.
        df_to_save = df.reindex(columns=config.EXPECTED_COLUMNS)
        df_to_save.to_pickle(config.STATUS_FILE)
        logging.info(f"Status data successfully saved to {config.STATUS_FILE}")
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        logging.error(f"Error saving status to {config.STATUS_FILE}: {e}", exc_info=True)
        messagebox.showerror("Error", f"Error saving status: {e}")


def process_data(new_df_raw: pd.DataFrame, current_status_df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Merges new Excel data with the current status DataFrame, handling new, 
    existing, and missing jobs.

    Args:
        new_df_raw (pd.DataFrame): Raw DataFrame loaded from Excel.
        current_status_df (pd.DataFrame): Current job status DataFrame.

    Returns:
        pd.DataFrame | None: The processed and merged DataFrame, or None if a critical
                             error (like missing 'Invoice #' after loading) occurs.
    """
    logging.info("Starting data processing: merging new Excel data with current status.")

    new_df_sanitized = new_df_raw.copy()
    new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}")

    # --- Critical Check: Ensure 'Invoice #' exists before merging ---
    if 'Invoice #' not in new_df_sanitized.columns:
        logging.error("Process_data: 'Invoice #' column is missing in the new data AFTER sanitization. Cannot proceed with merge.")
        # This message box might be redundant if load_excel already showed one, but serves as a final safeguard.
        messagebox.showerror("Processing Error", 
                             "The 'Invoice #' column could not be found in the loaded Excel data. "
                             "Please ensure the column exists and is correctly named.")
        return None # Stop processing if 'Invoice #' is missing

    # Ensure 'Invoice #' is treated as string to avoid merge issues if types differ
    new_df_sanitized['Invoice #'] = new_df_sanitized['Invoice #'].astype(str)
    current_status_df['Invoice #'] = current_status_df['Invoice #'].astype(str)

    # Drop potential '#' column if it exists and isn't expected
    if '#' in new_df_sanitized.columns and '#' not in config.EXPECTED_COLUMNS:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.")

    if '#' in current_status_df.columns and '#' not in config.EXPECTED_COLUMNS:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.")

    # --- Determine which columns from new_df_sanitized should be used for merging/updating ---
    # We want columns from new_df_sanitized that are *also* in EXPECTED_COLUMNS
    # but *exclude* 'Status' and 'Notes', as those should come from current_status_df.
    cols_from_new = [col for col in config.EXPECTED_COLUMNS 
                     if col in new_df_sanitized.columns and col not in ['Status', 'Notes']]
    
    # If 'Invoice #' wasn't in cols_from_new, add it. It's essential.
    if 'Invoice #' not in cols_from_new:
        cols_from_new.insert(0, 'Invoice #')

    logging.debug(f"Columns to use from new Excel for merge/update: {cols_from_new}")

    # Perform the merge using only the relevant columns from the new data.
    merged_df = pd.merge(
        current_status_df,
        new_df_sanitized[cols_from_new], # Use only selected columns
        on='Invoice #',
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}")

    processed_rows = []

    for index, row in merged_df.iterrows():
        current_row_data = {}
        invoice_num = row.get('Invoice #')

        if row['_merge'] == 'right_only': # Job is new (only in Excel)
            logging.debug(f"Processing new job (right_only): Invoice # {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                elif col == 'Status':
                    current_row_data[col] = 'New'
                elif col == 'Notes':
                    current_row_data[col] = ''
                else:
                    # New jobs get their data directly from the new columns (which don't have '_new' suffix here)
                    current_row_data[col] = row[col] if pd.notna(row.get(col)) else None

        elif row['_merge'] == 'left_only': # Job in current status, missing from new Excel
            logging.debug(f"Processing job missing from new Excel (left_only): Invoice # {invoice_num}")
            original_status_val = None
            existing_notes = ""
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                else:
                    # 'left_only' means data comes only from the _old side or columns that didn't merge
                    # Check for '_old' suffix first, then the original column name
                    old_col_name = col + '_old'
                    val = row[old_col_name] if old_col_name in row and pd.notna(row[old_col_name]) else (row[col] if col in row and pd.notna(row[col]) else None)
                    current_row_data[col] = val
                    if col == 'Status': original_status_val = val
                    if col == 'Notes': existing_notes = str(val) if val is not None else ""

            current_row_data['Status'] = config.REVIEW_MISSING_STATUS
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            alert_message = f"System Alert ({timestamp}): Job not in last Excel import."
            if original_status_val and original_status_val != config.REVIEW_MISSING_STATUS:
                alert_message += f" Previous status: '{original_status_val}'. "
            alert_message += "Verify if closed (set Status to 'Closed') or if active (re-include in Excel & update status)."
            current_row_data['Notes'] = (alert_message + "\n-----\n" + existing_notes).strip()
            logging.info(f"Invoice # {invoice_num}: Status set to '{config.REVIEW_MISSING_STATUS}' and Notes updated.")

        elif row['_merge'] == 'both': # Job exists in both
            logging.debug(f"Processing existing job (both): Invoice # {invoice_num}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = invoice_num
                elif col in ['Status', 'Notes']:
                    # For 'both', Status/Notes come from current_status_df (left side).
                    # Since they weren't in cols_from_new, they won't have a suffix.
                    current_row_data[col] = row[col] if pd.notna(row.get(col)) else ('' if col == 'Notes' else 'New')
                else:
                    # Other columns: Prioritize the new value, fall back to old if new is missing/NaN
                    new_val = row.get(col + '_new')
                    old_val = row.get(col + '_old') # This should exist since it came from current_status_df
                    current_row_data[col] = new_val if pd.notna(new_val) else old_val

        # Safeguard: Ensure all expected columns are present
        for col in config.EXPECTED_COLUMNS:
            if col not in current_row_data:
                logging.warning(f"Safeguard: Column '{col}' was missing for Invoice # {invoice_num}. Setting default.")
                current_row_data[col] = 'New' if col == 'Status' else ('' if col == 'Notes' else None)
        
        processed_rows.append(current_row_data)

    if not processed_rows:
        logging.info("No rows to process after merge. Returning empty DataFrame with expected columns.")
        final_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)
        logging.info(f"Successfully processed {len(final_df)} rows.")

    # Ensure final DataFrame matches the expected structure and order
    final_df = final_df.reindex(columns=config.EXPECTED_COLUMNS)
    # Ensure 'Invoice #' is still string before returning
    final_df['Invoice #'] = final_df['Invoice #'].astype(str)
    
    logging.info("Data processing finished.")
    return final_df