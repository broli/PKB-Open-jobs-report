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


def process_data(new_df_raw: pd.DataFrame, current_status_df: pd.DataFrame) -> pd.DataFrame:
    logging.info("Starting data processing: merging new Excel data with current status.")

    new_df_sanitized = new_df_raw.copy()
    new_df_sanitized.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in new_df_sanitized.columns]
    logging.debug(f"Sanitized new DataFrame columns: {list(new_df_sanitized.columns)}")

    if '#' in new_df_sanitized.columns and '#' not in config.EXPECTED_COLUMNS:
        new_df_sanitized = new_df_sanitized.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from new DataFrame as it's not in EXPECTED_COLUMNS.")

    if '#' in current_status_df.columns and '#' not in config.EXPECTED_COLUMNS:
        current_status_df = current_status_df.drop(columns=['#'], errors='ignore')
        logging.debug("Dropped '#' column from current_status_df as it's not in EXPECTED_COLUMNS.")

    merged_df = pd.merge(
        current_status_df,
        new_df_sanitized,
        on='Invoice #',
        how='outer',
        suffixes=('_old', '_new'),
        indicator=True
    )
    logging.debug(f"Merge completed. Merge indicator counts:\n{merged_df['_merge'].value_counts()}")

    processed_rows = []

    for index, row in merged_df.iterrows():
        current_row_data = {}

        if row['_merge'] == 'right_only': # Job is new (only in Excel)
            logging.debug(f"Processing new job (right_only): Invoice # {row.get('Invoice #')}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif col in new_df_sanitized.columns: # If this expected column was present in the new Excel
                    # Name in merged_df: col+'_new' if col was also in current_status_df, else 'col'
                    actual_name_in_merged = col + '_new' if (col + '_new') in row else col
                    val = row[actual_name_in_merged] if pd.notna(row[actual_name_in_merged]) else None
                    current_row_data[col] = val
                    if col == 'Status' and val is None: current_row_data[col] = 'New'
                    if col == 'Notes' and val is None: current_row_data[col] = ''
                else: # col is expected, but was NOT in new Excel (e.g. 'Status', 'Notes')
                    current_row_data[col] = 'New' if col == 'Status' else ('' if col == 'Notes' else None)

        elif row['_merge'] == 'left_only': # Job in current status, missing from new Excel
            logging.debug(f"Processing job missing from new Excel (left_only): Invoice # {row.get('Invoice #')}")
            original_status_val = None
            existing_notes = ""
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                else:
                    # Name in merged_df: col+'_old' if col was also in new_df_sanitized, else 'col'
                    actual_name_in_merged = col + '_old' if (col + '_old') in row else col
                    val = row[actual_name_in_merged] if pd.notna(row[actual_name_in_merged]) else None
                    current_row_data[col] = val
                    if col == 'Status' and val is not None: original_status_val = val
                    if col == 'Notes' and val is not None: existing_notes = str(val)
                    elif col == 'Notes' and val is None: current_row_data[col] = '' # Default notes to empty

            current_row_data['Status'] = config.REVIEW_MISSING_STATUS
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            alert_message = f"System Alert ({timestamp}): Job not in last Excel import."
            if original_status_val and original_status_val != config.REVIEW_MISSING_STATUS:
                alert_message += f" Previous status: '{original_status_val}'. "
            alert_message += "Verify if closed (set Status to 'Closed') or if active (re-include in Excel & update status)."
            current_row_data['Notes'] = (alert_message + "\n-----\n" + existing_notes).strip()
            logging.info(f"Invoice # {row.get('Invoice #')}: Status set to '{config.REVIEW_MISSING_STATUS}' and Notes updated.")

        elif row['_merge'] == 'both': # Job exists in both
            logging.debug(f"Processing existing job (both): Invoice # {row.get('Invoice #')}")
            for col in config.EXPECTED_COLUMNS:
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif col in ['Status', 'Notes']:
                    # These are from current_status_df (left DF).
                    # Since they are NOT in new_df_sanitized (right DF), no suffix is added. Name remains 'col'.
                    val = row[col] if pd.notna(row[col]) else None
                    current_row_data[col] = val if val is not None else ('' if col == 'Notes' else 'New')
                else: # Other columns (e.g. 'Account', 'Salesperson') that ARE in new_df_sanitized
                      # and also in current_status_df, so they get _new and _old suffixes.
                    new_val_col = col + '_new'
                    old_val_col = col + '_old'
                    if new_val_col in row and pd.notna(row[new_val_col]):
                        current_row_data[col] = row[new_val_col]
                    elif old_val_col in row and pd.notna(row[old_val_col]):
                        current_row_data[col] = row[old_val_col]
                    else: # Should not happen if col is in EXPECTED_COLUMNS and in new_df_sanitized
                        current_row_data[col] = None
            
            # Ensure Status/Notes have defaults if they were None from current_status_df
            if current_row_data.get('Status') is None: current_row_data['Status'] = 'New'
            if current_row_data.get('Notes') is None: current_row_data['Notes'] = ''

        # Safeguard: Ensure all expected columns are in current_row_data
        for col in config.EXPECTED_COLUMNS:
            if col not in current_row_data:
                logging.warning(f"Safeguard: Column '{col}' was missing for Invoice # {current_row_data.get('Invoice #', 'UNKNOWN')}. Setting default.")
                current_row_data[col] = 'New' if col == 'Status' else ('' if col == 'Notes' else None)
        
        processed_rows.append(current_row_data)

    if not processed_rows:
        logging.info("No rows to process after merge. Returning empty DataFrame with expected columns.")
        final_df = pd.DataFrame(columns=config.EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)
        logging.info(f"Successfully processed {len(final_df)} rows.")

    final_df = final_df.reindex(columns=config.EXPECTED_COLUMNS)
    logging.info("Data processing finished.")
    return final_df