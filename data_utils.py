import pandas as pd
from tkinter import filedialog, messagebox
import logging

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
OUTPUT_FILE = "open_invoices.xlsx"
EXPECTED_COLUMNS = [
    '#', 'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]
# DATE_FORMAT = '%b-%d'  # Moved to openJobs_class.py as it's UI related
# CURRENCY_COLUMNS = ['Invoice Total', 'Balance'] # Moved to openJobs_class.py as it's UI related
# CURRENCY_FORMAT = '${:,.2f}'  # Moved to openJobs_class.py as it's UI related

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
        return df
    except FileNotFoundError:
        # Create empty DataFrame if file doesn't exist
        logging.debug(f"File not found: {STATUS_FILE}, so creating empty DataFrame.")
        df = pd.DataFrame(columns=EXPECTED_COLUMNS)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error loading status: {e}")
        return None

def save_status(df):
    """Saves the current status DataFrame to a pickle file."""
    try:
        df.to_pickle(STATUS_FILE)
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving status: {e}")

def process_data(new_df, status_df):
    """
    Processes the new Excel data against the current status.
    """
    # 1. Sanitize column names in new_df
    new_df.columns = [str(col).strip() for col in new_df.columns]
    new_df.columns = [col.replace('\n', '').replace('\r', '') for col in new_df.columns]

    # 2. Merge DataFrames
    merged_df = new_df.merge(status_df, on='Invoice #', how='outer', indicator=True, suffixes=('_new', '_old'))

    def update_row(row):
        if row['_merge'] == 'both':
            for col in new_df.columns:
                if col != 'Invoice #':
                    new_col_name = col + '_new'
                    old_col_name = col + '_old'
                    if new_col_name in row:
                        row[col] = row[new_col_name]
                    elif old_col_name in row:
                        row[col] = row[old_col_name]
        elif row['_merge'] == 'left_only':
            for col in new_df.columns:
                if col != 'Invoice #':
                    new_col_name = col + '_new'
                    if new_col_name in row:
                        row[col] = row[new_col_name]
                    else:
                        row[col] = None
            row['Status'] = 'New'
            row['Notes'] = ''
        elif row['_merge'] == 'right_only':
            row = row.copy()
            row['Status'] = 'Closed'
        return row

    merged_df = merged_df.apply(update_row, axis=1)

    # 3. Clean up the merged DataFrame
    cols_to_drop = ['_merge']
    for col in new_df.columns:
        if col != 'Invoice #':
            cols_to_drop.append(col + '_new')
    for col in status_df.columns:
        if col not in ('Invoice #', 'Status', 'Notes'):
            cols_to_drop.append(col + '_old')

    final_df = merged_df.drop(columns=cols_to_drop, errors='ignore')

    # 4. Ensure final DataFrame has the expected columns
    for col in EXPECTED_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = None

    return final_df[EXPECTED_COLUMNS]