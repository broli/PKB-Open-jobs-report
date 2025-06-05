# config.py
# This file contains application-wide configurations and constants.

import logging
LOG_LEVEL = logging.DEBUG

# --- Application File Names ---
# OUTPUT_FILE = "open_invoices.xlsx"  # Default for generated Excel report (REMOVED as per plan)
STATUS_FILE = "job_data.db"    # For saving/loading current status (database file)
DB_TABLE_NAME = "jobs"         # Name of the table in the SQLite database

# --- Font Configuration ---
DEFAULT_FONT_FAMILY = 'Calibri'
DEFAULT_FONT_SIZE = 12
DEFAULT_FONT = (DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE)
DEFAULT_FONT_BOLD = (DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, 'bold')

# --- UI Layout & Formatting ---
DEFAULT_PADDING = 10
DATE_FORMAT = '%Y-%m-%d'  # Example: 2025-05-31 (UPDATED for YYYY-MM-DD format)
CURRENCY_COLUMNS = ['Invoice Total', 'Balance']  # Columns containing currency
CURRENCY_FORMAT = '${:,.2f}'  # Example: $1,234.56

# --- Data Specific Configuration ---
# Single source of truth for DataFrame column names and order
EXPECTED_COLUMNS = [
    'Invoice #', 'Order Date', 'Turn in Date', 'Account', # RETAINED 'Invoice #' based on your feedback
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]

# Predefined list of allowed statuses for an invoice
ALLOWED_STATUS = [
    "Waiting Measure", "Ready to order", "Waiting for materials",
    "Ready to dispatch", "In install", "Done", "Permit",
    "Cancelled/Postponed", "New", "Closed",
    "Review - Missing from Report"
]

# Specific status for jobs missing from a new report
REVIEW_MISSING_STATUS = "Review - Missing from Report"

# --- Treeview Column Widths ---
# Preferred initial widths for Treeview columns (in pixels)
PREFERRED_COLUMN_WIDTHS = {
    'Invoice #': 80,  # RETAINED 'Invoice #' based on your feedback
    'Order Date': 100,
    'Turn in Date': 100,
    'Account': 250,
    'Invoice Total': 100,
    'Balance': 100,
    'Salesperson': 120,
    'Project Coordinator': 150,
    'Status': 200,
    'Notes': 350
}
MIN_COLUMN_WIDTH = 20  # Minimum allowable width for any column
MAX_COLUMN_WIDTH = 500  # Maximum allowable width for any column

# --- Style Configuration for Treeview Rows ---
# Colors for different row statuses in the Treeview
STATUS_COLORS = {
    "default_fg": "black", "default_bg": "white",
    "action_needed_bg": "#FFEBEE", "action_needed_fg": "black",  # Light Red
    "all_good_bg": "#E8F5E9", "all_good_fg": "black",         # Light Green
    "closed_bg": "#E0E0E0", "closed_fg": "#757575",           # Light Grey
    "new_bg": "#E3F2FD", "new_fg": "black",                   # Light Blue
    "review_missing_bg": "#FFF9C4", "review_missing_fg": "#795548", # Light Yellow, Brown text
    "selected_bg": "#B0BEC5", "selected_fg": "black"          # Blue Grey for selected rows
}

# --- Logging Configuration ---
LOG_LEVEL = logging.DEBUG  # DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'

# --- Reporting Tab UI ---
REPORT_SUB_TAB_BG_COLOR = "#F0F8FF"
REPORT_TEXT_FG_COLOR = "black"

# --- Application Information (Optional) ---
APP_NAME = "Open Jobs Status Tracker"
APP_VERSION = "2.0.1" # UPDATED version for migration