import pandas as pd
from tkinter import filedialog, messagebox
import logging

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
OUTPUT_FILE = "open_invoices.xlsx" # This is not directly used in this file but often kept with related constants
EXPECTED_COLUMNS = [
    'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]


def load_excel(excel_file):
    """Loads an Excel file into a Pandas DataFrame."""
    try:
        df = pd.read_excel(excel_file)
        return df
    except FileNotFoundError:
        messagebox.showerror("Error", f"File not found: {excel_file}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel: {e}")
        return None

def load_status():
    """Loads the current status DataFrame from a pickle file."""
    try:
        logging.debug(f"Trying to load status from {STATUS_FILE}")
        df = pd.read_pickle(STATUS_FILE)
        # Ensure loaded data has all expected columns, adding missing ones with None
        for col in EXPECTED_COLUMNS:
            if col not in df.columns:
                df[col] = None
        return df[EXPECTED_COLUMNS] # Ensure correct column order and selection
    except FileNotFoundError:
        # Create empty DataFrame with expected columns if file doesn't exist
        logging.debug(f"File not found: {STATUS_FILE}, so creating empty DataFrame with predefined columns.")
        df = pd.DataFrame(columns=EXPECTED_COLUMNS)
        return df
    except Exception as e:
        logging.error(f"Error loading status: {e}. Creating empty DataFrame.")
        messagebox.showerror("Error", f"Error loading status: {e}. A new empty dataset will be used.")
        # Return an empty DataFrame with the correct structure in case of other errors
        return pd.DataFrame(columns=EXPECTED_COLUMNS)


def save_status(df):
    """Saves the current status DataFrame to a pickle file."""
    try:
        # Ensure only expected columns are saved
        df_to_save = df[EXPECTED_COLUMNS].copy()
        df_to_save.to_pickle(STATUS_FILE)
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving status: {e}")

def process_data(new_df, status_df):
    """
    Processes the new Excel data against the current status.
    Ensures the '#' column is not part of the processing if it exists in new_df.
    """
    # 1. Sanitize column names in new_df
    new_df.columns = [str(col).strip() for col in new_df.columns]
    new_df.columns = [col.replace('\n', '').replace('\r', '') for col in new_df.columns]

    # Ensure '#' column is dropped from new_df if it exists, before any merging.
    if '#' in new_df.columns:
        new_df = new_df.drop(columns=['#'], errors='ignore')

    # Ensure status_df also doesn't have it, though it should be controlled by EXPECTED_COLUMNS
    if '#' in status_df.columns:
        status_df = status_df.drop(columns=['#'], errors='ignore')

    # 2. Merge DataFrames
    # We use 'Invoice #' as the key. Other columns from new_df will update existing ones.
    # Columns in status_df not in new_df (like 'Status', 'Notes' if not in Excel) should be preserved.

    # Identify columns that are in new_df for the update (excluding the merge key)
    update_cols_from_new = [col for col in new_df.columns if col != 'Invoice #']

    # Merge, giving preference to new data for common columns, except for 'Status' and 'Notes' initially
    # which are handled specially.
    merged_df = pd.merge(status_df, new_df, on='Invoice #', how='outer', suffixes=('_old', '_new'), indicator=True)

    # Initialize 'Status' and 'Notes' for new rows
    # For 'left_only' (only in old status_df), mark as 'Closed' unless already handled.
    # For 'right_only' (only in new new_df), mark as 'New'.
    # For 'both', preserve old status and notes unless explicitly overwritten.

    processed_rows = []
    for index, row in merged_df.iterrows():
        # Create a dictionary for the new row to build upon
        current_row_data = {}

        if row['_merge'] == 'right_only': # New invoice
            for col in EXPECTED_COLUMNS:
                new_col_name = col + '_new' # Data comes from the new_df side
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif new_col_name in row and pd.notna(row[new_col_name]):
                    current_row_data[col] = row[new_col_name]
                elif col == 'Status':
                    current_row_data[col] = 'New'
                elif col == 'Notes':
                    current_row_data[col] = ''
                else:
                    current_row_data[col] = None # Or an appropriate default
            # Fill any EXPECTED_COLUMNS not in new_df with defaults
            for col in EXPECTED_COLUMNS:
                if col not in current_row_data:
                    if col == 'Status': current_row_data[col] = 'New'
                    elif col == 'Notes': current_row_data[col] = ''
                    else: current_row_data[col] = None


        elif row['_merge'] == 'left_only': # Invoice only in old status, potentially closed
            all_old_data_present = True
            for col in EXPECTED_COLUMNS:
                old_col_name = col + '_old'
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif old_col_name in row and pd.notna(row[old_col_name]):
                    current_row_data[col] = row[old_col_name]
                elif col in row and pd.notna(row[col]): # If column name didn't have suffix
                    current_row_data[col] = row[col]
                else:
                    # If essential data is missing from a 'left_only' it might be an issue
                    # For now, we assume it's being marked closed.
                    current_row_data[col] = None # Or retain if already exists
                    if col not in ['Status', 'Notes']: all_old_data_present = False


            # If it was truly only in the old data and not just an artifact of no new data for it
            # we mark it as 'Closed'. However, status_df should ideally already manage its statuses.
            # This logic handles cases where an invoice is removed from the input Excel.
            if 'Status' not in current_row_data or pd.isna(current_row_data['Status']):
                 current_row_data['Status'] = 'Closed' # Default for items no longer in new Excel
            if 'Notes' not in current_row_data: current_row_data['Notes'] = current_row_data.get('Notes_old', '')


        elif row['_merge'] == 'both': # Invoice in both old status and new Excel
            for col in EXPECTED_COLUMNS:
                new_col_name = col + '_new'
                old_col_name = col + '_old'
                if col == 'Invoice #':
                    current_row_data[col] = row[col]
                elif col in ['Status', 'Notes']: # Preserve old Status/Notes unless new data overwrites (which it won't by default with this merge)
                    current_row_data[col] = row[old_col_name] if pd.notna(row[old_col_name]) else (row[new_col_name] if new_col_name in row and pd.notna(row[new_col_name]) else None)
                    if col == 'Status' and pd.isna(current_row_data[col]): current_row_data[col] = 'New' # if old status was NaN
                    if col == 'Notes' and pd.isna(current_row_data[col]): current_row_data[col] = ''
                elif new_col_name in row and pd.notna(row[new_col_name]): # Data from new Excel takes precedence for other fields
                    current_row_data[col] = row[new_col_name]
                elif old_col_name in row and pd.notna(row[old_col_name]): # Fallback to old data if not in new
                    current_row_data[col] = row[old_col_name]
                else:
                    current_row_data[col] = None

        processed_rows.append(current_row_data)

    if not processed_rows:
        final_df = pd.DataFrame(columns=EXPECTED_COLUMNS)
    else:
        final_df = pd.DataFrame(processed_rows)

    # Ensure final DataFrame has exactly the EXPECTED_COLUMNS in the correct order
    # Add any missing EXPECTED_COLUMNS (shouldn't happen with above logic but good failsafe)
    for col in EXPECTED_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = None # Or appropriate default like '' for Notes, 'New' for Status if it's a truly new column

    return final_df[EXPECTED_COLUMNS]