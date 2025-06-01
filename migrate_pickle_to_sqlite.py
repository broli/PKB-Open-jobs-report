# migrate_pickle_to_sqlite.py
import pandas as pd
import sqlite3
import os
import datetime # For Timestamp.now() in date adjustment

# Attempt to import config from the current directory
try:
    import config
except ImportError:
    print("Error: config.py not found. Make sure this script is in the same directory as config.py.")
    exit()

# --- Configuration ---
OLD_PICKLE_FILE = "invoice_status.pkl" # Assumes it's in the same directory
NEW_DB_FILE = config.STATUS_FILE       # From config.py (e.g., "job_data.db")
DB_TABLE_NAME = config.DB_TABLE_NAME   # From config.py (e.g., "jobs")
DATE_COLUMNS_TO_CONVERT = ['Order Date', 'Turn in Date']

# --- Helper function for date year adjustment (adapted from data_utils.py) ---
def adjust_ambiguous_date_years(date_series: pd.Series, current_timestamp: pd.Timestamp, series_name: str = "Unknown") -> pd.Series:
    """
    Adjusts years for dates in a Series that might have been ambiguously parsed.
    If a date was parsed into the current year but is later than the current date,
    it's assumed to be from the previous year.
    """
    if not isinstance(date_series, pd.Series) or date_series.empty or not pd.api.types.is_datetime64_any_dtype(date_series):
        if isinstance(date_series, pd.Series) and not date_series.empty:
            print(f"Info: Series '{series_name}' is not of datetime type or is empty. Skipping year adjustment.")
        return date_series

    adjusted_series = date_series.copy()
    # Mask for dates in the current year but later than the current date
    mask = (adjusted_series.dt.year == current_timestamp.year) & \
           (adjusted_series > current_timestamp) & \
           (adjusted_series.notna())

    if mask.any():
        adjusted_series.loc[mask] = adjusted_series.loc[mask] - pd.DateOffset(years=1)
        print(f"Info: Adjusted year for some dates in series '{series_name}' assuming they were from the previous year.")
    
    return adjusted_series

def migrate_data():
    print(f"Starting data migration from '{OLD_PICKLE_FILE}' to SQLite database '{NEW_DB_FILE}', table '{DB_TABLE_NAME}'.")

    # 1. Check if old pickle file exists
    if not os.path.exists(OLD_PICKLE_FILE):
        print(f"Error: Old pickle file '{OLD_PICKLE_FILE}' not found. Migration aborted.")
        return

    try:
        # 2. Load data from pickle file
        print(f"Loading data from '{OLD_PICKLE_FILE}'...")
        df = pd.read_pickle(OLD_PICKLE_FILE)
        print(f"Successfully loaded {len(df)} rows from pickle file.")

        # 3. Data Transformation (Important: Keep 'Invoice #' as per user requirement)
        # No column renaming for 'Invoice #' to 'PO #'

        # 4. Date Conversion and Adjustment
        print("Converting and adjusting date columns...")
        current_ts = pd.Timestamp.now()
        for col in DATE_COLUMNS_TO_CONVERT:
            if col in df.columns:
                print(f"  Processing date column: '{col}'")
                # Convert to datetime, attempting to infer format. 
                # If dates are strings like 'May-17', to_datetime might need a format hint
                # or handle it well if they are already datetime objects in pickle.
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = adjust_ambiguous_date_years(df[col], current_ts, series_name=col)
            else:
                print(f"  Warning: Date column '{col}' not found in the pickle data.")
        
        # 5. Ensure DataFrame structure matches EXPECTED_COLUMNS from config
        # This is good practice to catch any discrepancies before writing to SQL.
        # Add missing columns with appropriate defaults (None or NaT for dates)
        for expected_col in config.EXPECTED_COLUMNS:
            if expected_col not in df.columns:
                print(f"Warning: Column '{expected_col}' not found in pickle. Adding it with default values.")
                if expected_col in DATE_COLUMNS_TO_CONVERT:
                    df[expected_col] = pd.NaT
                else:
                    df[expected_col] = None
        
        # Reorder columns to match config.EXPECTED_COLUMNS
        df = df.reindex(columns=config.EXPECTED_COLUMNS)
        print("DataFrame columns reordered and validated against config.EXPECTED_COLUMNS.")

        # 6. Connect to SQLite database (creates the file if it doesn't exist)
        print(f"Connecting to SQLite database '{NEW_DB_FILE}'...")
        conn = sqlite3.connect(NEW_DB_FILE)
        
        # 7. Save DataFrame to SQLite table
        print(f"Saving data to table '{DB_TABLE_NAME}' (replacing if exists)...")
        # Using if_exists='replace' will drop the table first if it exists and then create a new one.
        df.to_sql(DB_TABLE_NAME, conn, if_exists='replace', index=False)
        print("Data saved successfully to SQLite.")

    except FileNotFoundError:
        print(f"Error: Pickle file '{OLD_PICKLE_FILE}' not found.")
    except pd.errors.EmptyDataError:
        print(f"Error: Pickle file '{OLD_PICKLE_FILE}' is empty.")
    except sqlite3.Error as e_sql:
        print(f"SQLite error during migration: {e_sql}")
    except Exception as e:
        print(f"An unexpected error occurred during migration: {e}")
    finally:
        if 'conn' in locals() and conn:
            conn.close()
            print("SQLite connection closed.")

if __name__ == "__main__":
    migrate_data()