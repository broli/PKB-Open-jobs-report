import pandas as pd
from tkinter import messagebox # filedialog is not used here, messagebox is.
import logging
import datetime # For timestamping the alert note

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
# EXPECTED_COLUMNS defines the standard structure for the DataFrame.
# It's the single source of truth for column names and their order.
EXPECTED_COLUMNS = [
    'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]

# Define the specific status to use when a job is missing from the report
REVIEW_MISSING_STATUS = "Review - Missing from Report"


def load_excel(excel_file_path: str) -> pd.DataFrame | None:
    """
    Loads data from an Excel file into a Pandas DataFrame.

    Args:
        excel_file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame | None: A DataFrame containing the Excel data,
                             or None if an error occurs during loading.
    """
    try:
        logging.debug(f"Attempting to load Excel file from: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        logging.info(f"Successfully loaded Excel file: {excel_file_path}")
        return df
    except FileNotFoundError:
        logging.error(f"Excel file not found at path: {excel_file_path}")
        messagebox.showerror("Error", f"File not found: {excel_file_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading Excel file {excel_file_path}: {e}", exc_info=True)
        messagebox.showerror("Error", f"Error reading Excel: {e}")
        return None


def load_status() -> pd.DataFrame:
    """
    Loads the current job status DataFrame from a pickle file (STATUS_FILE).

    If the file is not found, or if there's an error loading it,
    an empty DataFrame with the EXPECTED_COLUMNS structure is returned.
    Ensures that the loaded DataFrame conforms to EXPECTED_COLUMNS,
    adding any missing columns with None values and ensuring correct order.

    Returns:
        pd.DataFrame: The loaded (or newly created) status DataFrame.
    """
    try:
        logging.debug(f"Trying to load status from {STATUS_FILE}")
        df = pd.read_pickle(STATUS_FILE)
        logging.info(f"Successfully loaded status data from {STATUS_FILE}")

        # Ensure loaded data has all expected columns, adding missing ones with None
        # This maintains schema consistency.
        for col in EXPECTED_COLUMNS:
            if col not in df.columns:
                logging.warning(f"Column '{col}' missing in loaded status data. Adding it with None values.")
                df[col] = None
        # Return DataFrame with columns in the defined order and only expected columns
        return df[EXPECTED_COLUMNS]
    except FileNotFoundError:
        logging.info(f"Status file '{STATUS_FILE}' not found. Creating a new empty DataFrame.")
        return pd.DataFrame(columns=EXPECTED_COLUMNS)
    except Exception as e:
        logging.error(f"Error loading status from {STATUS_FILE}: {e}. Creating empty DataFrame.", exc_info=True)
        messagebox.showerror("Error", f"Error loading status data: {e}. A new empty dataset will be used.")
        return pd.DataFrame(columns=EXPECTED_COLUMNS)


def save_status(df: pd.DataFrame) -> None:
    """
    Saves the current status DataFrame to a pickle file (STATUS_FILE).

    Only the EXPECTED_COLUMNS are saved, ensuring a consistent schema.

    Args:
        df (pd.DataFrame): The DataFrame containing the current job statuses to save.
    """
    if df is None:
        logging.warning("Attempted to save a None DataFrame. Operation skipped.")
        messagebox.showwarning("Save Warning", "No data to save.")
        return
    try:
        # Ensure only expected columns are saved, in the correct order.
        df_to_save = df[EXPECTED_COLUMNS].copy()
        df_to_save.to_pickle(STATUS_FILE)
        logging.info(f"Status data successfully saved to {STATUS_FILE}")
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        logging.error(f"Error saving status to {STATUS_FILE}: {e}", exc_info=True)
        messagebox.showerror("Error", f"Error saving status: {e}")


def process_data(new_df_raw: pd.DataFrame, current_status_df: pd.DataFrame) -> pd.DataFrame:
    """
    Processes new job data from an Excel import against the current status data.

    The function performs the following key operations:
    1.  Sanitizes column names in the new raw DataFrame.
    2.  Merges the new data with the current status data using 'Invoice #' as the key.
    3.  Identifies new jobs, existing jobs, and jobs present in current status but missing from the new import.
    4.  For new jobs ('right_only' in merge): Assigns 'New' status.
    5.  For existing jobs ('both' in merge): Updates data from the new import, preserving existing 'Status' and 'Notes'.
    6.  For jobs missing from new import ('left_only' in merge):
        - Preserves their data from `current_status_df`.
        - Changes their 'Status' to `REVIEW_MISSING_STATUS` (e.g., "Review - Missing from Report").
        - Prepends an alert to their 'Notes', including their original status and a timestamp.
    7.  Ensures the final DataFrame strictly adheres to `EXPECTED_COLUMNS` structure and order.

    Args:
        new_df_raw (pd.DataFrame): The raw DataFrame loaded from the new Excel file.
        current_status_df (pd.DataFrame): The DataFrame holding the current job statuses.

    Returns:
        pd.DataFrame: The processed DataFrame with updated job statuses and information.
    """
    logging.info("Starting data processing: merging new Excel data with current status.")

    # --- 1. Sanitize column names in the new DataFrame ---
    # Ensures consistency by stripping whitespace and removing newline characters.
    new_df_sanitized = new_df_raw.copy()
    new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}")

    # Drop the common incidental '#' column if it exists in the new data.
    if '#' in new_df_sanitized.columns:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from new DataFrame.")

    # Ensure current_status_df also doesn't have it (should be controlled by EXPECTED_COLUMNS anyway)
    if '#' in current_status_df.columns:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from current status DataFrame.")

    # --- 2. Merge DataFrames ---
    # 'outer' merge includes all jobs from both DataFrames.
    # `indicator=True` adds a '_merge' column ('left_only', 'right_only', 'both').
    # Suffixes help differentiate columns with the same name from old and new DataFrames.
    merged_df = pd.merge(
        current_status_df,
        new_df_sanitized,
        on='Invoice #',
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}")

    processed_rows = [] # List to hold dictionaries, each representing a processed row.

    # --- 3. Process each row based on merge status ---
    for index, row in merged_df.iterrows():
        current_row_data = {} # To build the data for the final DataFrame row.

        if row['_merge'] == 'right_only':
            # --- Case: New job (only in new Excel data) ---
            logging.debug(f"Processing new job (right_only): Invoice # {row.get('Invoice #')}")
            for col in EXPECTED_COLUMNS:
                new_col_data_field = col + '_new' # Field name from new data after merge
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif new_col_data_field in row and pd.notna(row[new_col_data_field]):
                    current_row_data[col] = row[new_col_data_field]
                elif col == 'Status':
                    current_row_data[col] = 'New' # Default status for new jobs
                elif col == 'Notes':
                    current_row_data[col] = '' # Default empty notes for new jobs
                else:
                    current_row_data[col] = None # Default other fields to None
            # Ensure all EXPECTED_COLUMNS are present even if not in new_df_sanitized
            for col in EXPECTED_COLUMNS:
                if col not in current_row_data:
                    current_row_data[col] = 'New' if col == 'Status' else ('' if col == 'Notes' else None)

        elif row['_merge'] == 'left_only':
            # --- Case: Job in current status, but NOT in new Excel (potentially closed or needs review) ---
            logging.debug(f"Processing job missing from new Excel (left_only): Invoice # {row.get('Invoice #')}")
            original_status_val = None
            existing_notes = ""

            # Populate data from the old status DataFrame first
            for col in EXPECTED_COLUMNS:
                old_col_data_field = col + '_old' # Field name from old data after merge
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif old_col_data_field in row and pd.notna(row[old_col_data_field]):
                    current_row_data[col] = row[old_col_data_field]
                    if col == 'Status':
                        original_status_val = row[old_col_data_field]
                    if col == 'Notes':
                        existing_notes = str(row[old_col_data_field]) # Ensure notes are string
                elif col in row and pd.notna(row[col]): # Fallback for 'Invoice #' itself or if no suffix (should not happen for _old)
                    current_row_data[col] = row[col]
                else: # Column was not in old data or was NaN
                    current_row_data[col] = '' if col == 'Notes' else None # Default notes to empty, others to None

            # Set the status to indicate it needs review
            current_row_data['Status'] = REVIEW_MISSING_STATUS

            # Construct and prepend an alert message to the Notes
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            alert_message = f"System Alert ({timestamp}): Job not in last Excel import."
            if original_status_val and original_status_val != REVIEW_MISSING_STATUS: # Avoid "Previous status was 'Review...'"
                alert_message += f" Previous status: '{original_status_val}'. "
            alert_message += "Verify if closed (then set Status to 'Closed') or if still active (re-include in Excel & update status)."

            # Combine alert with existing notes
            current_row_data['Notes'] = (alert_message + "\n-----\n" + existing_notes).strip()
            logging.info(f"Invoice # {row.get('Invoice #')}: Status set to '{REVIEW_MISSING_STATUS}' and Notes updated with alert.")

        elif row['_merge'] == 'both':
            # --- Case: Existing job (in both current status and new Excel) ---
            logging.debug(f"Processing existing job (both): Invoice # {row.get('Invoice #')}")
            for col in EXPECTED_COLUMNS:
                new_col_data_field = col + '_new'
                old_col_data_field = col + '_old'

                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif col in ['Status', 'Notes']:
                    # Preserve existing Status and Notes from old data, as these are managed by the user in-app.
                    current_row_data[col] = row[old_col_data_field] if pd.notna(row[old_col_data_field]) else \
                                          ('' if col == 'Notes' else 'New') # Default if old status/notes were NaN
                elif new_col_data_field in row and pd.notna(row[new_col_data_field]):
                    # For other fields, data from new Excel takes precedence.
                    current_row_data[col] = row[new_col_data_field]
                elif old_col_data_field in row and pd.notna(row[old_col_data_field]):
                    # Fallback to old data if not present in new Excel (e.g., new Excel has fewer columns).
                    current_row_data[col] = row[old_col_data_field]
                else:
                    # If data is not in new or old (e.g. a newly added EXPECTED_COLUMN), set to None.
                    current_row_data[col] = None
            # If status was somehow NaN from old data, default it
            if pd.isna(current_row_data.get('Status')):
                current_row_data['Status'] = 'New'
            if pd.isna(current_row_data.get('Notes')):
                current_row_data['Notes'] = ''


        # Ensure all EXPECTED_COLUMNS are present in current_row_data before appending
        for col in EXPECTED_COLUMNS:
            if col not in current_row_data:
                # This is a fallback, ideally all columns are handled above.
                logging.warning(f"Column '{col}' was missing from current_row_data for Invoice # {current_row_data.get('Invoice #')}. Setting to default.")
                if col == 'Status': current_row_data[col] = 'New'
                elif col == 'Notes': current_row_data[col] = ''
                else: current_row_data[col] = None
        processed_rows.append(current_row_data)

    # --- 4. Create Final DataFrame ---
    if not processed_rows:
        logging.info("No rows to process after merge. Returning empty DataFrame with expected columns.")
        final_df = pd.DataFrame(columns=EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)
        logging.info(f"Successfully processed {len(final_df)} rows.")

    # Final check to ensure the DataFrame structure (column existence and order).
    for col in EXPECTED_COLUMNS:
        if col not in final_df.columns:
            logging.warning(f"Final DataFrame is missing expected column '{col}'. Adding it with None values.")
            final_df[col] = None
    final_df = final_df[EXPECTED_COLUMNS] # Enforce column order and selection

    logging.info("Data processing finished.")
    return final_df