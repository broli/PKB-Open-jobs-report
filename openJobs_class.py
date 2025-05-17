# Import necessary tkinter modules for GUI creation
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar # ttk for themed widgets

# Import the Sun Valley ttk theme for a modern look
import sv_ttk

# Import pandas for data manipulation, particularly with DataFrames
import pandas as pd

# Import functions and constants from the data_utils.py file
# EXPECTED_COLUMNS is the single source of truth for column names and order
from data_utils import (
    load_status,
    save_status,
    load_excel,
    process_data,
    EXPECTED_COLUMNS
)

# Import the logging module for recording events and debugging
import logging

# --- Application Configuration Variables ---

# Default filename for the generated Excel report of open invoices
OUTPUT_FILE = "open_invoices.xlsx"

# --- Font Configuration ---
# Define a consistent font family and size for the application
DEFAULT_FONT_FAMILY = 'Calibri'
DEFAULT_FONT_SIZE = 12
# Standard font tuple (family, size)
DEFAULT_FONT = (DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE)
# Bold font tuple (family, size, style)
DEFAULT_FONT_BOLD = (DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, 'bold')

# --- UI Layout & Formatting ---
# Default padding around widgets
DEFAULT_PADDING = 10
# Date format string for displaying dates in the Treeview
DATE_FORMAT = '%b-%d' # Example: May-17
# List of column names that contain currency values
CURRENCY_COLUMNS = ['Invoice Total', 'Balance']
# Format string for displaying currency values
CURRENCY_FORMAT = '${:,.2f}' # Example: $1,234.56

# --- Data Specific Configuration ---
# Predefined list of allowed statuses for an invoice
ALLOWED_STATUS = [
    "Waiting Measure", "Ready to order", "Waiting for materials",
    "Ready to dispatch", "In install", "Done", "Permit",
    "Cancelled/Postponed", "New", "Closed"
]

# Dictionary defining preferred initial widths for Treeview columns (in pixels)
# These provide a better default layout than automatic calculation for all columns.
PREFERRED_COLUMN_WIDTHS = {
    'Invoice #': 80,
    'Order Date': 100,
    'Turn in Date': 100,
    'Account': 250,
    'Invoice Total': 100,
    'Balance': 100,
    'Salesperson': 120,
    'Project Coordinator': 150,
    'Status': 170,
    'Notes': 350
}

# Minimum and maximum allowable width for any column in the Treeview
MIN_COLUMN_WIDTH = 20
MAX_COLUMN_WIDTH = 500

# Configure basic logging for the application
# DEBUG level will show all logs; for production, might change to INFO or WARNING
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')

class OpenJobsApp(tk.Tk):
    """
    Main application class for the Open Jobs Status Tracker.
    Inherits from tk.Tk to create the main application window.
    Manages the UI, data display, and interactions for tracking job statuses.
    """
    def __init__(self):
        """
        Initialize the OpenJobsApp.
        Sets up the main window, theme, global font settings, styles,
        loads initial data, and creates the UI widgets.
        """
        super().__init__() # Initialize the base tk.Tk class

        # Apply the Sun Valley Ttk theme (light or dark)
        sv_ttk.set_theme("light")
        self.title("Open Jobs Status Tracker") # Set the window title

        # --- Global Font Configuration (Applied BEFORE WIDGET CREATION) ---
        # Uses Tkinter's option database to set default fonts for various widget types.
        # This is a powerful way to achieve consistent font styling.
        default_font_string = f"{DEFAULT_FONT_FAMILY} {DEFAULT_FONT_SIZE}"
        self.option_add("*Font", default_font_string) # General fallback font
        self.option_add("*Text*Font", default_font_string) # For tk.Text widgets
        self.option_add("*Label*Font", default_font_string) # For tk.Label (if used)
        self.option_add("*Button*Font", default_font_string) # For tk.Button (if used)
        self.option_add("*Menu*Font", default_font_string) # For tk.Menu and its items
        self.option_add("*MenuItem*Font", default_font_string) # More specific for menu items
        # For the dropdown list part of a ttk.Combobox (can be theme/OS dependent)
        self.option_add("*TCombobox*Listbox*Font", default_font_string)

        # Attempt to maximize the application window on startup
        self.maximize_window()

        # --- TTK Style Configuration ---
        # Create a ttk.Style object to customize the appearance of ttk widgets
        self.style = ttk.Style(self)

        # Set the default font for ALL ttk widgets by configuring the root style '.'
        # This font will be inherited by specific ttk widgets unless overridden.
        self.style.configure('.', font=DEFAULT_FONT)

        # Configure specific ttk widget styles
        # Treeview: Set font and row height for better readability
        self.style.configure("Treeview", font=DEFAULT_FONT, rowheight=30)
        # Treeview Headings: Set a bold font
        self.style.configure("Treeview.Heading", font=DEFAULT_FONT_BOLD)
        # TCombobox: Ensure it uses the default font (though '.' should cover it)
        self.style.configure("TCombobox", font=DEFAULT_FONT)

        # Define colors for different row statuses in the Treeview
        self.status_colors = {
            "default_fg": "black", "default_bg": "white", # Default text and background
            "action_needed_bg": "#FFEBEE", "action_needed_fg": "black", # Light Red background
            "all_good_bg": "#E8F5E9", "all_good_fg": "black",        # Light Green background
            "closed_bg": "#E0E0E0", "closed_fg": "#757575",          # Light Grey bg, Darker Grey text
            "new_bg": "#E3F2FD", "new_fg": "black",                  # Light Blue background
            "selected_bg": "#B0BEC5", "selected_fg": "black"         # Blue Grey for selected rows
        }

        # Map the 'selected' state of Treeview rows to specific background and foreground colors
        self.style.map("Treeview",
                       background=[('selected', self.status_colors["selected_bg"])],
                       foreground=[('selected', self.status_colors["selected_fg"])])

        # --- Data Loading ---
        # Load the initial status data from the pickle file (via data_utils.py)
        self.status_df = load_status()

        # Handle cases where data loading might fail or return None
        if self.status_df is None:
            messagebox.showerror("Fatal Error", "Could not load or initialize status data. Exiting.")
            self.destroy() # Close the application
            return

        # Ensure all EXPECTED_COLUMNS are present in the loaded DataFrame.
        # load_status should handle this, but this is a fallback.
        missing_cols = [col for col in EXPECTED_COLUMNS if col not in self.status_df.columns]
        if missing_cols:
            messagebox.showwarning("Data Warning",
                                   f"Loaded data is missing columns: {', '.join(missing_cols)}. "
                                   f"They will be added with default values (None).")
            for col in missing_cols:
                self.status_df[col] = None # Add missing columns with None

        # Ensure the DataFrame uses exactly the EXPECTED_COLUMNS in the correct order
        self.status_df = self.status_df[EXPECTED_COLUMNS]

        # --- UI Creation ---
        # Create all the main UI widgets (menu, treeview, scrollbars)
        self.create_widgets()
        # Variable to keep track of any active popup editing window
        self.editing_window = None

        # Populate the Treeview with the loaded data
        self.populate_treeview()

    def maximize_window(self):
        """Attempts to maximize the application window to fill the screen."""
        try:
            self.state('zoomed') # Works on Windows
        except tk.TclError: # Fallback for other systems (some Linux/macOS)
            try:
                # Get maximum screen dimensions and set geometry
                m = self.maxsize()
                self.geometry('{}x{}+0+0'.format(*m))
            except tk.TclError: # Further fallback to fullscreen attribute
                self.attributes('-fullscreen', True)

    def create_widgets(self):
        """Creates and configures the main UI widgets of the application."""
        # --- Menu Bar ---
        # tk.Menu font should be influenced by `self.option_add("*Menu*Font", ...)`
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar, tearoff=0) # tearoff=0 removes the dashed line
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel)
        file_menu.add_command(label="Save Current Status", command=self.save_data)
        file_menu.add_command(label="Generate Open Jobs Report", command=self.generate_report)
        file_menu.add_separator() # Adds a dividing line in the menu
        file_menu.add_command(label="Exit", command=self.quit_app)
        menu_bar.add_cascade(label="File", menu=file_menu) # Adds the "File" menu to the menu bar
        self.config(menu=menu_bar) # Set the menu bar for the main window

        # --- Treeview (Main Data Display) ---
        # Create the Treeview widget with columns defined by EXPECTED_COLUMNS
        # show="headings" hides the default first empty column ('#0')
        # selectmode='extended' allows multiple rows to be selected
        self.tree = ttk.Treeview(self, columns=EXPECTED_COLUMNS, show="headings", selectmode='extended')

        # Configure each column heading and properties
        for col in EXPECTED_COLUMNS:
            # Set column heading text and make it clickable for sorting
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            # Set initial column width from PREFERRED_COLUMN_WIDTHS, with a default
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W) # tk.W = West (left-align)

        # --- Scrollbars for Treeview ---
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview) # Vertical scrollbar
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview) # Horizontal scrollbar
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set) # Link scrollbars to treeview

        # Pack (place) the scrollbars and treeview in the window
        vsb.pack(side=tk.RIGHT, fill=tk.Y) # Vertical scrollbar on the right, fills vertically
        hsb.pack(side=tk.BOTTOM, fill=tk.X) # Horizontal scrollbar at the bottom, fills horizontally
        self.tree.pack(fill=tk.BOTH, expand=True, padx=DEFAULT_PADDING, pady=DEFAULT_PADDING) # Treeview fills remaining space

        # --- Event Bindings for Treeview ---
        self.tree.bind("<Double-1>", self.on_double_click) # Double-click to edit Status/Notes
        self.tree.bind("<Delete>", self.delete_selected_row) # Delete key to remove selected row(s)


    def configure_treeview_columns(self):
        """
        Ensures the Treeview columns match the current DataFrame's columns.
        This is typically called if the DataFrame structure might change,
        though with EXPECTED_COLUMNS, it should remain consistent.
        """
        current_tree_cols = list(self.status_df.columns)
        self.tree.configure(columns=current_tree_cols)
        for col in current_tree_cols:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)
        # Ensure the default '#0' column (if it somehow appears) is hidden
        if '#0' in self.tree.column('#0'):
            self.tree.column('#0', width=0, stretch=tk.NO)

    def populate_treeview(self):
        """
        Clears and repopulates the Treeview with data from self.status_df.
        Formats dates and currency values for display.
        Assigns DataFrame index as a tag to each Treeview item for later reference.
        """
        # Clear any existing items from the Treeview
        for i in self.tree.get_children():
            self.tree.delete(i)

        # If no data, log and return
        if self.status_df is None or self.status_df.empty:
            logging.info("No data to populate in the treeview.")
            return

        # Ensure DataFrame columns are in the EXPECTED_COLUMNS order
        self.status_df = self.status_df[EXPECTED_COLUMNS]

        date_columns = ['Order Date', 'Turn in Date'] # Columns to format as dates

        # Iterate over DataFrame rows and insert them into the Treeview
        for df_index, row in self.status_df.iterrows(): # df_index is the actual index in the DataFrame
            values = [] # List to hold formatted values for the current row
            for col_name in EXPECTED_COLUMNS: # Iterate in the defined column order
                value = row[col_name]
                # Format date columns
                if col_name in date_columns and pd.notna(value):
                    try:
                        value = pd.to_datetime(value).strftime(DATE_FORMAT)
                    except (ValueError, TypeError):
                        logging.warning(f"Could not format date for '{value}' in column '{col_name}'. Original value used.")
                # Format currency columns
                elif col_name in CURRENCY_COLUMNS and pd.notna(value):
                    try:
                        num_value = float(str(value).replace('$', '').replace(',', '')) # Clean and convert to float
                        value = CURRENCY_FORMAT.format(num_value)
                    except (ValueError, TypeError):
                        logging.warning(f"Could not format currency for '{value}' in column '{col_name}'. Original value used.")
                # Append the (possibly formatted) value, or an empty string if NaN/None
                values.append(value if pd.notna(value) else "")

            # Insert the row into the Treeview.
            # The first tag is the DataFrame index (as a string), used to link Treeview item to DataFrame row.
            self.tree.insert("", tk.END, values=tuple(values), tags=(str(df_index),))

        # Schedule column width adjustment and row coloring shortly after population
        self.after(10, self.set_column_widths_from_preferred) # Use preferred widths
        self.color_rows() # Apply status-based coloring

    def load_new_excel(self):
        """
        Handles the "Load New Excel" menu command.
        Opens a file dialog to select an Excel file, loads it,
        processes it against the current status data, and updates the Treeview.
        """
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if not excel_file_path: # User cancelled dialog
            return

        new_df_raw = load_excel(excel_file_path) # From data_utils
        if new_df_raw is None: # Error handled in load_excel
            return

        # Make a copy of the current status DataFrame for processing
        current_status_df_copy = self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=EXPECTED_COLUMNS)

        # Process the new Excel data against the current status (from data_utils)
        # process_data handles merging, new/closed statuses, and uses EXPECTED_COLUMNS.
        self.status_df = process_data(new_df_raw, current_status_df_copy)

        if self.status_df is None: # Should not happen if process_data is robust
            messagebox.showerror("Error", "Failed to process the new Excel data. Reverting to previous data.")
            self.status_df = current_status_df_copy # Revert to the copy before processing
            return

        # Ensure final DataFrame conforms to EXPECTED_COLUMNS structure and order
        self.status_df = self.status_df[EXPECTED_COLUMNS]

        self.configure_treeview_columns() # Reconfigure tree (though columns shouldn't change)
        self.populate_treeview() # Refresh the display with new data
        messagebox.showinfo("Success", "New Excel data loaded and processed.")

    def save_data(self):
        """
        Handles the "Save Current Status" menu command.
        Saves the current self.status_df to a pickle file using data_utils.save_status.
        """
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.")
            return
        save_status(self.status_df) # From data_utils

    def generate_report(self):
        """
        Handles the "Generate Open Jobs Report" menu command.
        Filters for open invoices, prompts for a save location,
        and exports the filtered data to an Excel file with formatting.
        """
        if self.status_df is None or self.status_df.empty:
            messagebox.showinfo("Info", "No data available to generate a report.")
            return

        # Filter DataFrame for open invoices (Status is not 'Closed' or 'Cancelled/Postponed')
        open_invoices_df = self.status_df[
            ~self.status_df['Status'].isin(['Closed', 'Cancelled/Postponed'])
        ].copy() # .copy() to avoid SettingWithCopyWarning

        # Ensure only EXPECTED_COLUMNS are in the report, in the correct order
        open_invoices_df = open_invoices_df[EXPECTED_COLUMNS]

        if open_invoices_df.empty:
            messagebox.showinfo("Info", "No open invoices to report.")
            return

        try:
            # Ask user where to save the report
            report_file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Open Invoices Report As",
                initialfile=OUTPUT_FILE # Default filename
            )
            if not report_file_path: # User cancelled
                messagebox.showinfo("Info", "Report generation cancelled.")
                return

            # Use pandas.ExcelWriter with xlsxwriter engine for formatting
            with pd.ExcelWriter(report_file_path, engine='xlsxwriter') as writer:
                open_invoices_df.to_excel(writer, index=False, sheet_name='Open Invoices')
                workbook = writer.book
                worksheet = writer.sheets['Open Invoices']

                # Apply currency formatting to relevant columns in the Excel sheet
                currency_format_excel = workbook.add_format({'num_format': '$#,##0.00'})
                for col_name in CURRENCY_COLUMNS:
                    if col_name in open_invoices_df.columns:
                        col_idx = open_invoices_df.columns.get_loc(col_name)
                        col_letter = chr(ord('A') + col_idx) # Convert 0-based index to Excel column letter
                        # Apply format to the entire column (None means no specific width)
                        worksheet.set_column(f'{col_letter}:{col_letter}', None, currency_format_excel)

                # Auto-adjust column widths in Excel based on content
                for i, col in enumerate(open_invoices_df.columns):
                    # Max length of data in column or header, plus a little padding
                    column_len = max(open_invoices_df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len)

            messagebox.showinfo("Success", f"Report generated successfully: {report_file_path}")
        except Exception as e:
            logging.error(f"Error generating report: {e}", exc_info=True) # Log full traceback
            messagebox.showerror("Error", f"An error occurred while generating the report: {e}")

    def on_double_click(self, event):
        """
        Handles double-click events on Treeview rows.
        Opens an editor for 'Status' or 'Notes' columns.
        """
        item_id = self.tree.identify_row(event.y) # Get Treeview item ID under mouse
        column_id_str = self.tree.identify_column(event.x) # Get Treeview column ID (e.g., "#1")

        if not item_id or not column_id_str: # Click was not on a valid item/column
            return

        try:
            # Convert tree column ID (e.g., "#1") to 0-indexed integer
            column_index_tree = int(column_id_str.replace("#", "")) - 1

            # Validate column index against EXPECTED_COLUMNS
            if not (0 <= column_index_tree < len(EXPECTED_COLUMNS)):
                logging.warning(f"Invalid column index from tree: {column_index_tree}")
                return

            # Get the actual column name using the index and EXPECTED_COLUMNS
            actual_column_name = EXPECTED_COLUMNS[column_index_tree]

            # Retrieve the DataFrame index stored as the first tag of the Treeview item
            df_row_index_str = self.tree.item(item_id, "tags")[0]
            df_row_index = int(df_row_index_str) # This is the direct index in self.status_df

            # Close any existing editor window before opening a new one
            if self.editing_window and self.editing_window.winfo_exists():
                self.editing_window.destroy()
            self.editing_window = None

            # Open specific editor based on the column name
            if actual_column_name == "Status":
                self.create_status_editor(item_id, df_row_index, actual_column_name)
            elif actual_column_name == "Notes":
                self.create_notes_editor(item_id, df_row_index, actual_column_name)
            else:
                # No special editor for other columns on double-click
                logging.debug(f"No special editor for column '{actual_column_name}'. Double-click ignored.")
        except (ValueError, IndexError, TypeError) as e:
            logging.error(f"Error in on_double_click: {e}. Item ID: {item_id}, Column ID Str: {column_id_str}", exc_info=True)

    def _common_editor_save(self, item_id, df_row_index, column_name, new_value, editor_window):
        """
        Helper function to save edited data from popup editors.
        Updates the DataFrame (self.status_df) and the Treeview display.
        Closes the editor window and re-colors rows if necessary.

        Args:
            item_id: The ID of the Treeview item being edited.
            df_row_index: The integer index of the row in self.status_df.
            column_name: The name of the column being edited.
            new_value: The new value to save.
            editor_window: The Toplevel editor window to be closed.
        """
        try:
            # Update the DataFrame using .loc for precise row/column targeting
            self.status_df.loc[df_row_index, column_name] = new_value

            # Update the corresponding cell in the Treeview
            current_tree_values = list(self.tree.item(item_id, "values"))
            if column_name in EXPECTED_COLUMNS:
                column_tree_idx = EXPECTED_COLUMNS.index(column_name) # Get 0-based index for tree values list

                # Reformat the value for display if it's a currency or date column
                display_value = new_value
                if column_name in CURRENCY_COLUMNS and pd.notna(new_value):
                    try: display_value = CURRENCY_FORMAT.format(float(str(new_value).replace('$', '').replace(',', '')))
                    except: pass # If formatting fails, use raw new_value
                elif column_name in ['Order Date', 'Turn in Date'] and pd.notna(new_value):
                    try: display_value = pd.to_datetime(new_value).strftime(DATE_FORMAT)
                    except: pass # If formatting fails, use raw new_value

                current_tree_values[column_tree_idx] = display_value
                self.tree.item(item_id, values=tuple(current_tree_values)) # Update tree item
            else:
                logging.error(f"Column '{column_name}' not in EXPECTED_COLUMNS during save. Treeview not updated for this column.")

            editor_window.destroy() # Close the editor popup
            self.editing_window = None # Reset tracker
            self.color_rows() # Re-apply row coloring as status might have changed
        except Exception as e:
            logging.error(f"Error in _common_editor_save for column '{column_name}': {e}", exc_info=True)
            messagebox.showerror("Save Error", f"Could not save change for {column_name}: {e}")

    def create_notes_editor(self, item_id, df_row_index, actual_column_name):
        """
        Creates a Toplevel window with a tk.Text widget to edit the 'Notes' field.
        """
        self.editing_window = tk.Toplevel(self) # Create a new top-level window
        self.editing_window.title(f"Edit {actual_column_name}")
        self.editing_window.transient(self) # Keep it on top of the main window
        self.editing_window.grab_set() # Make it modal (block interaction with main window)

        # tk.Text font is influenced by `self.option_add("*Text*Font", ...)` or can be set explicitly.
        text_widget = tk.Text(self.editing_window, width=60, height=10, wrap=tk.WORD, font=DEFAULT_FONT)
        current_value = str(self.status_df.loc[df_row_index, actual_column_name]) # Get current note
        text_widget.insert(tk.END, current_value) # Populate text widget
        text_widget.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING, fill=tk.BOTH, expand=True)
        text_widget.focus() # Set focus to the text widget

        # Frame for Save/Cancel buttons
        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(0,DEFAULT_PADDING), padx=DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)

        # ttk.Button font is influenced by `self.style.configure('.', font=DEFAULT_FONT)`
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton", # Accent style for emphasis
                              command=lambda: self._common_editor_save(item_id, df_row_index, actual_column_name,
                                                                      text_widget.get("1.0", tk.END).strip(), # Get all text, strip whitespace
                                                                      self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)

        # Handle window close button (the 'X')
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.center_toplevel(self.editing_window) # Center the popup
        self.wait_window(self.editing_window) # Wait for the editor window to close

    def create_status_editor(self, item_id, df_row_index, column_name):
        """
        Creates a Toplevel window with a ttk.Combobox to edit the 'Status' field.
        """
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self); self.editing_window.grab_set()

        current_value = str(self.status_df.loc[df_row_index, column_name])
        status_var = StringVar(self.editing_window) # StringVar to link to Combobox
        if current_value in ALLOWED_STATUS:
            status_var.set(current_value)
        elif ALLOWED_STATUS: # Default to first allowed status if current is invalid
            status_var.set(ALLOWED_STATUS[0])
        else: # Should not happen if ALLOWED_STATUS is populated
            status_var.set("")

        # Display Invoice # for context
        inv_num_col = 'Invoice #'
        inv_num = self.status_df.loc[df_row_index, inv_num_col] if inv_num_col in self.status_df.columns else "N/A"
        # ttk.Label font is influenced by `self.style.configure('.', font=DEFAULT_FONT)`
        ttk.Label(self.editing_window, text=f"Status for Invoice {inv_num}:").pack(padx=DEFAULT_PADDING,pady=(DEFAULT_PADDING,5))

        # ttk.Combobox entry font is styled by `self.style.configure("TCombobox", ...)`
        # Dropdown list font is influenced by `self.option_add("*TCombobox*Listbox*Font", ...)`
        combo = ttk.Combobox(self.editing_window, textvariable=status_var, values=ALLOWED_STATUS, state="readonly") # Readonly prevents typing
        combo.pack(padx=DEFAULT_PADDING, pady=5, fill=tk.X); combo.focus()

        # Frame for Save/Cancel buttons
        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(5,DEFAULT_PADDING), padx=DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._common_editor_save(item_id, df_row_index, column_name,
                                                                      status_var.get(), self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)

        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.center_toplevel(self.editing_window)
        self.wait_window(self.editing_window)

    def delete_selected_row(self, event=None):
        """
        Handles the <Delete> key press. Deletes selected row(s) from the
        DataFrame and refreshes the Treeview.
        """
        selected_tree_items = self.tree.selection() # Get IDs of selected Treeview items
        if not selected_tree_items:
            messagebox.showinfo("No Selection", "Please select one or more rows to delete.")
            return

        confirm_msg = f"Are you sure you want to delete {len(selected_tree_items)} selected row(s)? This action cannot be undone from the UI."
        if not messagebox.askyesno("Confirm Delete", confirm_msg):
            return

        # Get DataFrame indices from the tags of selected Treeview items
        # Sort in reverse to avoid index shifting issues when deleting multiple rows
        df_indices_to_delete = sorted([int(self.tree.item(item_id, "tags")[0]) for item_id in selected_tree_items], reverse=True)

        if df_indices_to_delete:
            self.status_df.drop(index=df_indices_to_delete, inplace=True) # Delete rows from DataFrame
            self.status_df.reset_index(drop=True, inplace=True) # Reset DataFrame index to be contiguous (0, 1, 2...)

        # Repopulate the treeview entirely to reflect deletions and new DataFrame indices.
        # This ensures Treeview tags (DataFrame indices) are always correct.
        self.populate_treeview()
        messagebox.showinfo("Success", f"{len(selected_tree_items)} row(s) deleted.")

    def set_column_widths_from_preferred(self):
        """Sets Treeview column widths based on the PREFERRED_COLUMN_WIDTHS dictionary."""
        self.update_idletasks() # Ensure UI is up-to-date before getting/setting widths
        for col_name in EXPECTED_COLUMNS:
            if col_name == '#0': continue # Skip the hidden internal column
            width = PREFERRED_COLUMN_WIDTHS.get(col_name, 100) # Get preferred width or default to 100
            final_width = max(MIN_COLUMN_WIDTH, int(width)) # Ensure min width
            final_width = min(MAX_COLUMN_WIDTH, int(final_width)) # Ensure max width
            self.tree.column(col_name, width=final_width, anchor=tk.W)
        # Ensure the #0 column (if it exists and is shown by mistake) is hidden
        if '#0' in self.tree['columns']:
            self.tree.column('#0', width=0, stretch=tk.NO)

    def color_rows(self):
        """
        Applies background and foreground colors to Treeview rows based on their 'Status'.
        Uses pre-configured tag styles.
        """
        if self.status_df is None or self.status_df.empty:
            return

        # Define styles (tags) for different statuses using the configured colors
        styles_map = {
            "default_status_style": (self.status_colors["default_bg"], self.status_colors["default_fg"]),
            "action_needed_style": (self.status_colors["action_needed_bg"], self.status_colors["action_needed_fg"]),
            "all_good_style": (self.status_colors["all_good_bg"], self.status_colors["all_good_fg"]),
            "closed_style": (self.status_colors["closed_bg"], self.status_colors["closed_fg"]),
            "new_style": (self.status_colors["new_bg"], self.status_colors["new_fg"])
        }
        # Configure each tag in the Treeview with its background and foreground
        for tag_name, (bg, fg) in styles_map.items():
            self.tree.tag_configure(tag_name, background=bg, foreground=fg)

        # Define which statuses fall into which color category
        action_statuses = ["Ready to order", "Permit", "Waiting Measure"]
        good_statuses = ["Ready to dispatch", "In install", "Done", "Waiting for materials"]
        # 'New', 'Closed', 'Cancelled/Postponed' have their own specific styles

        status_col_name = "Status" # The column name containing the status
        if status_col_name not in EXPECTED_COLUMNS:
            logging.error(f"'{status_col_name}' column missing from EXPECTED_COLUMNS. Cannot color rows accurately.")
            return

        try:
            # Get the 0-based index of the 'Status' column in the Treeview display
            status_column_tree_index = EXPECTED_COLUMNS.index(status_col_name)
        except ValueError:
            logging.error(f"'{status_col_name}' column not found in Treeview columns (EXPECTED_COLUMNS). Cannot color rows.")
            return

        # Iterate over all items (rows) in the Treeview
        for item_id in self.tree.get_children():
            try:
                # The first tag of the item is the DataFrame index string
                df_index_tag = self.tree.item(item_id, "tags")[0]
                new_tags_for_item = [df_index_tag] # Start with the index tag

                values = self.tree.item(item_id, "values") # Get all cell values for the row
                # Ensure values exist and the status column index is valid
                if values and len(values) > status_column_tree_index:
                    status = str(values[status_column_tree_index]) # Get the status string

                    # Apply the appropriate style tag based on the status value
                    if status == "New":
                        new_tags_for_item.append("new_style")
                    elif status == "Closed" or status == "Cancelled/Postponed":
                        new_tags_for_item.append("closed_style")
                    elif status in action_statuses:
                        new_tags_for_item.append("action_needed_style")
                    elif status in good_statuses:
                        new_tags_for_item.append("all_good_style")
                    else: # Default style for any other status
                        new_tags_for_item.append("default_status_style")
                else: # Fallback if status somehow not found for the row
                    new_tags_for_item.append("default_status_style")

                # Apply the collected tags to the Treeview item
                self.tree.item(item_id, tags=tuple(new_tags_for_item))
            except Exception as e:
                logging.error(f"Error coloring row {item_id}: {e}", exc_info=True)
                # Apply default styling if an error occurs for a specific row
                df_idx_tag_fallback = self.tree.item(item_id,"tags")[0] if self.tree.item(item_id,"tags") else "err_idx_fallback"
                self.tree.item(item_id, tags=(df_idx_tag_fallback, "default_status_style"))

    def sort_treeview_column(self, col, reverse):
        """
        Sorts the Treeview rows based on the clicked column.
        Attempts to sort numerically or by date if possible, otherwise alphabetically.

        Args:
            col: The name of the column to sort by.
            reverse: Boolean, True for descending sort, False for ascending.
        """
        if self.status_df is None or self.status_df.empty:
            return

        try:
            # Get data from Treeview for sorting. Each item is (value_in_column, item_id).
            # Sorting based on displayed values in Treeview.
            data_list = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]

            # Define a sort key function to handle different data types
            def sort_key_func(item):
                val_str = item[0] # Value from tree (always a string initially)
                # Attempt to convert to numeric for sorting
                # Handle currency by stripping '$' and ','
                if isinstance(val_str, str) and val_str.startswith('$'):
                    try: return float(val_str.replace('$', '').replace(',', ''))
                    except ValueError: return val_str.lower() # Fallback to string sort
                # Handle dates (assuming DATE_FORMAT is sortable or convert back)
                try: return pd.to_datetime(val_str, format=DATE_FORMAT) # Try parsing with app's date format
                except (ValueError, TypeError): pass # If not a date in this format, try next
                # Handle general numbers
                try: return float(val_str)
                except (ValueError, TypeError): return str(val_str).lower() # Fallback to case-insensitive string sort

            data_list.sort(key=sort_key_func, reverse=reverse)

            # Reorder items in the Treeview according to the sorted list
            for index, (val, k) in enumerate(data_list):
                self.tree.move(k, '', index) # Move item k to the new index

            # Update the heading command to toggle sort direction for the next click
            self.tree.heading(col, command=lambda _col=col: self.sort_treeview_column(_col, not reverse))
        except Exception as e:
            logging.error(f"Error sorting column {col}: {e}", exc_info=True)

    def center_toplevel(self, toplevel_window):
        """
        Centers a Toplevel window (popup) relative to the main application window.
        """
        toplevel_window.update_idletasks() # Ensure window dimensions are calculated

        # Main window geometry
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()

        # Toplevel window geometry
        pop_width = toplevel_window.winfo_width()
        pop_height = toplevel_window.winfo_height()

        # Calculate position for the Toplevel window to be centered
        x = main_x + (main_width // 2) - (pop_width // 2)
        y = main_y + (main_height // 2) - (pop_height // 2)

        toplevel_window.geometry(f"+{x}+{y}") # Set the Toplevel window's position

    def quit_app(self):
        """
        Handles the "Exit" menu command.
        Prompts the user to save changes before quitting.
        """
        if messagebox.askokcancel("Quit", "Do you want to save changes before quitting?"):
            self.save_data() # Save data if user confirms
        self.destroy() # Close the main application window


# This block allows the script to be run directly for testing the OpenJobsApp class.
# When imported as a module, this block will not execute.
if __name__ == '__main__':
    app = OpenJobsApp() # Create an instance of the application
    app.mainloop()      # Start the Tkinter event loop
