# app_shell.py
# This module defines the main application class for the Open Jobs Status Tracker.
# It sets up the main window, manages data, and integrates different UI tabs.

# --- Standard Library Imports ---
import tkinter as tk  # For creating the GUI
from tkinter import filedialog, messagebox, ttk  # Specific Tkinter components
import logging # For logging application events and errors

# --- Third-Party Library Imports ---
import pandas as pd  # For data manipulation, primarily with DataFrames
import sv_ttk  # For applying a modern theme to Tkinter widgets

# --- Local Application Imports ---
import config  # Stores application-wide configurations and constants
import data_utils  # Contains utility functions for data loading, saving, and processing

# Import Tab classes that represent different sections of the UI
from data_management_tab import DataManagementTab  # Tab for managing and editing job data
from reporting_tab import ReportingTab  # Tab for displaying statistics and reports
from export_tab import ExportTab # <<< IMPORTED ExportTab

# --- Logging Configuration ---
# Set up basic logging for the application.
# The log level and format are defined in the 'config' module.
# This should be done once, early in the application's lifecycle.
logging.basicConfig(level=config.LOG_LEVEL, format=config.LOG_FORMAT)

class OpenJobsApp(tk.Tk):
    """
    Main application class for the Open Jobs Status Tracker.
    Inherits from tk.Tk to create the main application window.
    It manages the overall UI structure (menu, notebook for tabs),
    handles core data operations (loading, saving, processing),
    and facilitates communication between different parts of the application.
    """
    def __init__(self):
        """
        Initializes the OpenJobsApp.
        Sets up the main window, theme, fonts, loads initial data,
        and creates the main UI layout.
        """
        super().__init__()  # Call the constructor of the parent class (tk.Tk)

        # --- Theme and Window Setup ---
        sv_ttk.set_theme("light")  # Apply the 'light' theme from sv_ttk
        self.title(config.APP_NAME + " - v" + config.APP_VERSION)  # Set the window title
        # Define a custom action for the window's close button (WM_DELETE_WINDOW)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

        # --- Application Configuration Access ---
        # Make common configuration values from the 'config' module
        # easily accessible as attributes of the app instance.
        self.DEFAULT_FONT = config.DEFAULT_FONT
        self.DEFAULT_FONT_BOLD = config.DEFAULT_FONT_BOLD
        self.CURRENCY_FORMAT = config.CURRENCY_FORMAT
        # Note: EXPECTED_COLUMNS is primarily used by config, data_utils,
        # and data_management_tab directly. Tabs can import 'config'
        # if they need specific constants.

        # --- Global Font Configuration ---
        # Set a default font for various Tkinter widget types.
        default_font_string = f"{config.DEFAULT_FONT_FAMILY} {config.DEFAULT_FONT_SIZE}"
        self.option_add("*Font", default_font_string)
        self.option_add("*Text*Font", default_font_string)
        self.option_add("*Label*Font", default_font_string)
        self.option_add("*Button*Font", default_font_string)
        self.option_add("*Menu*Font", default_font_string)
        self.option_add("*MenuItem*Font", default_font_string)
        self.option_add("*TCombobox*Listbox*Font", default_font_string)

        self.maximize_window()  # Maximize the application window on startup

        # --- Style Configuration ---
        # Configure global and specific widget styles.
        self.style = ttk.Style(self)
        self.style.configure('.', font=config.DEFAULT_FONT)  # Global style for all ttk widgets
        # Configure the appearance of selected items in Treeview widgets.
        # Colors are sourced from the 'config' module.
        self.style.map("Treeview",
                       background=[('selected', config.STATUS_COLORS["selected_bg"])],
                       foreground=[('selected', config.STATUS_COLORS["selected_fg"])])

        # --- Main Application Data ---
        # self.status_df will hold the primary DataFrame for job statuses.
        self.status_df = None  # Initialize as None
        self.load_initial_data()  # Load existing data or create an empty structure

        # --- UI Creation ---
        # Initialize UI components that will be created later.
        self.notebook = None  # The main ttk.Notebook widget for tabs
        self.data_tab_instance = None  # Instance of the DataManagementTab
        self.reporting_tab_instance = None  # Instance of the ReportingTab
        self.export_tab_instance = None # <<< INITIALIZED export_tab_instance
        
        self.create_main_ui_layout()  # Create the main UI elements (menu, notebook, tabs)
        
        # --- Initial Data Population for UI ---
        # If the data management tab exists and data is loaded, populate its treeview.
        if self.data_tab_instance and self.status_df is not None:
            self.data_tab_instance.populate_treeview()

    def load_initial_data(self):
        """
        Loads initial status data from a persistent file (e.g., a pickle file)
        or creates an empty DataFrame if no data exists or an error occurs.
        Ensures the loaded DataFrame conforms to the expected column structure.
        """
        # Load status data using a utility function.
        self.status_df = data_utils.load_status()
        
        # Handle cases where data loading might fail or return None.
        if self.status_df is None:
            messagebox.showerror("Fatal Error", "Could not load or initialize status data. Exiting.")
            self.destroy()  # Close the application if critical data cannot be loaded
            return
            
        # Ensure the DataFrame has all columns defined in config.EXPECTED_COLUMNS.
        # Add any missing columns and fill them with None.
        missing_cols = [col for col in config.EXPECTED_COLUMNS if col not in self.status_df.columns]
        if missing_cols:
            logging.warning(f"AppShell: Loaded data is missing columns: {', '.join(missing_cols)}. Adding them.")
            for col in missing_cols:
                self.status_df[col] = None  # Add missing column with default None values
        # Reindex to ensure the columns are in the order defined in config.EXPECTED_COLUMNS.
        self.status_df = self.status_df.reindex(columns=config.EXPECTED_COLUMNS)

    def maximize_window(self):
        """
        Maximizes the application window to fill the screen.
        Tries different methods for cross-platform compatibility.
        """
        try:
            self.state('zoomed')  # Works on Windows
        except tk.TclError:
            try:
                # Attempt to get max size and set geometry (common on some Linux WMs)
                m = self.maxsize()
                self.geometry('{}x{}+0+0'.format(*m))
            except tk.TclError:
                # Fallback to fullscreen attribute (can be less ideal than maximized)
                self.attributes('-fullscreen', True)

    def create_main_ui_layout(self):
        """
        Creates the main UI structure of the application, including the
        menu bar and the notebook widget that will contain different tabs.
        """
        # --- Menu Bar Creation ---
        menu_bar = tk.Menu(self)  # Create the main menu bar

        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)  # Create a "File" menu
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel_data)
        file_menu.add_command(label="Save Current Status", command=self.save_current_data)
        file_menu.add_command(label="Generate Open Jobs Report (Excel)", command=self.generate_excel_report)
        file_menu.add_separator()  # Add a visual separator line
        file_menu.add_command(label="Exit", command=self.quit_app)
        menu_bar.add_cascade(label="File", menu=file_menu)  # Add "File" menu to the menu bar

        self.config(menu=menu_bar)  # Set the menu bar for the main window

        # --- Notebook for Tabs ---
        # Create a ttk.Notebook widget to manage different application tabs.
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=config.DEFAULT_PADDING, pady=(0, config.DEFAULT_PADDING))

        # Create and Add Data Management Tab
        self.data_tab_instance = DataManagementTab(self.notebook, self) # 'self' (app instance) is passed to the tab
        self.notebook.add(self.data_tab_instance, text='Data Management')

        # Create and Add Reporting Tab
        self.reporting_tab_instance = ReportingTab(self.notebook, self) # 'self' (app instance) is passed to the tab
        self.notebook.add(self.reporting_tab_instance, text='Reporting & Statistics')
        
        # Create and Add Export Tab 
        self.export_tab_instance = ExportTab(self.notebook, self) # <<< INSTANTIATED ExportTab
        self.notebook.add(self.export_tab_instance, text='Export Report') # <<< ADDED ExportTab TO NOTEBOOK
        
        # Bind an event handler for when the selected notebook tab changes.
        self.notebook.bind("<<NotebookTabChanged>>", self.on_notebook_tab_changed)

    def on_notebook_tab_changed(self, event=None):
        """
        Event handler called when the active tab in the notebook changes.
        It notifies the newly selected tab (if the tab implements an
        'on_tab_selected' method).
        Args:
            event: The event object (optional, usually not directly used).
        """
        if not self.notebook:  # Safety check
            return
        try:
            # Get the widget corresponding to the currently selected tab.
            selected_tab_widget = self.notebook.nametowidget(self.notebook.select())
            logging.debug(f"AppShell: Tab changed to: {self.notebook.tab(self.notebook.select(), 'text')}")
            
            # If the selected tab has an 'on_tab_selected' method, call it.
            # This allows tabs to refresh or update themselves when they become active.
            if hasattr(selected_tab_widget, 'on_tab_selected'):
                selected_tab_widget.on_tab_selected()
        except tk.TclError:
            # Handle potential error if the tab widget cannot be identified.
            logging.warning("AppShell: Error identifying selected tab during NotebookTabChanged event.")

    def notify_data_changed(self):
        """
        Notifies relevant parts of the application (e.g., the reporting tab)
        that the underlying data (self.status_df) has changed.
        This allows other components to update their views or calculations.
        """
        logging.debug("AppShell: Data changed, notifying reporting tab.")
        # If the reporting tab instance exists and has an 'on_tab_selected' method,
        # call it to trigger a refresh. This is a general way to update the tab.
        # A more specific 'refresh_data' method could also be implemented on tabs.
        if self.reporting_tab_instance and hasattr(self.reporting_tab_instance, 'on_tab_selected'):
            self.reporting_tab_instance.on_tab_selected() 
            # Consider if DataManagementTab also needs a specific notification
            # beyond repopulating its own tree (which it typically handles itself).
        
        # Notify ExportTab as well, if it exists and has an on_tab_selected or similar method
        if self.export_tab_instance and hasattr(self.export_tab_instance, 'on_tab_selected'): # Or 'on_data_changed'
             self.export_tab_instance.on_tab_selected() # Call its on_tab_selected if it exists
             # For Phase 1, this is not strictly necessary to call a specific on_data_changed,
             # but on_tab_selected is good practice for all tabs.
             pass


    # --- Data Operation Methods ---
    def load_new_excel_data(self):
        """
        Handles the process of loading new job data from an Excel file.
        It prompts the user to select a file, loads the data, processes it
        against the current status, updates the main DataFrame, and refreshes the UI.
        Includes error handling for missing columns (KeyError).
        """
        # Open a file dialog to let the user select an Excel file.
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")],
            parent=self  # Set the main window as the parent of the dialog
        )
        if not excel_file_path:  # User cancelled the dialog
            return

        # Load the raw data from the selected Excel file.
        new_df_raw = data_utils.load_excel(excel_file_path)
        if new_df_raw is None:
            # data_utils.load_excel() shows its own error message (e.g., FileNotFoundError).
            return

        # Create a copy of the current status DataFrame to merge with new data.
        # If no current data, start with an empty DataFrame matching the expected structure.
        current_status_df_copy = self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=config.EXPECTED_COLUMNS)
        
        processed_df = None  # Initialize to ensure it's defined in all paths
        try:
            # Process the new raw data against the current status data.
            # This is the primary function where a KeyError might occur if new_df_raw
            # is missing critical columns that process_data expects, or if process_data
            # itself encounters an issue related to column names.
            processed_df = data_utils.process_data(new_df_raw, current_status_df_copy)

            if processed_df is None:
                # This case handles if process_data returns None due to an internal, unhandled issue.
                messagebox.showerror("Processing Error",
                                     "Failed to process the new Excel data. The data might be invalid or incomplete.",
                                     parent=self)
                logging.error("AppShell: data_utils.process_data returned None.")
                return

        except KeyError as e:
            # Catch KeyError specifically, which usually indicates a missing or misnamed column.
            # Display an informative error message to the user.
            
            # Construct a simple string of expected column names separated by " | "
            expected_columns_display = " | ".join(config.EXPECTED_COLUMNS) #

            error_message = (
                f"We encountered an unexpected situation with the Excel file that prevents us from proceeding.\n\n"
                f"The issue appears to be that the application could not find an expected column named: '{e}'.\n\n"
                f"The application requires specific column headers in your Excel sheet, in a particular order. "
                f"Please check the following in your Excel file:\n"
                f"  - All required columns are present in the first row of your data.\n"
                f"  - Column names exactly match the expected names (they are case-sensitive).\n"
                f"  - There are no merged cells in the header row (the row with column names).\n"
                f"  - The data starts directly under the header row. Avoid extra title rows, logos, or empty rows above the actual column headers.\n"
                f"  - Ensure there are no hidden characters or leading/trailing spaces in the column names within the Excel file cells.\n\n"
                f"Expected column names (in order):\n"
                f"{expected_columns_display}\n\n"
                f"The operation has been cancelled. Please correct the Excel file's header row and try again."
            )
            messagebox.showerror("Excel Format Issue", error_message, parent=self)
            logging.error(f"AppShell: KeyError during Excel data processing: Column '{e}' not found or other column mismatch.", exc_info=True)
            return  # Stop further processing to prevent errors with incomplete data

        except Exception as e:
            # Catch any other unexpected errors that might occur during data processing.
            messagebox.showerror("Processing Error",
                                 f"An unexpected error occurred while processing the Excel data: {e}",
                                 parent=self)
            logging.error(f"AppShell: Unexpected error during Excel data processing: {e}", exc_info=True)
            return
        
        # If processing was successful, update the main status DataFrame.
        # Reindex to ensure the DataFrame conforms to the expected schema.
        self.status_df = processed_df.reindex(columns=config.EXPECTED_COLUMNS)

        # Refresh the UI elements that display this data.
        if self.data_tab_instance:
            self.data_tab_instance.populate_treeview()  # Update the treeview in the Data Management tab
        messagebox.showinfo("Success", "New Excel data loaded and processed successfully.", parent=self)
        self.notify_data_changed()  # Notify other parts of the app (e.g., reporting tab)

    def save_current_data(self):
        """
        Saves the current state of the status_df to a persistent file
        (typically a pickle file, as defined in config.STATUS_FILE).
        """
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.", parent=self)
            return
        # The data_utils.save_status function handles the actual saving and shows its own success/error messages.
        data_utils.save_status(self.status_df)

    def generate_excel_report(self):
        """
        Generates an Excel report of open jobs.
        Filters out closed/cancelled jobs, prompts the user for a save location,
        and writes the report to an .xlsx file with formatting.
        """
        if self.status_df is None or self.status_df.empty:
            messagebox.showinfo("Info", "No data available to generate a report.", parent=self)
            return

        # Define statuses to exclude from the "Open Jobs" report.
        excluded_statuses = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
        # Filter the DataFrame to include only open jobs.
        open_invoices_df = self.status_df[~self.status_df['Status'].isin(excluded_statuses)].copy()
        # Ensure the report DataFrame has the correct column structure.
        open_invoices_df = open_invoices_df.reindex(columns=config.EXPECTED_COLUMNS)

        if open_invoices_df.empty:
            messagebox.showinfo("Info", "No open invoices to report.", parent=self)
            return

        try:
            # Ask the user where to save the report file.
            report_file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Open Invoices Report As",
                initialfile=config.OUTPUT_FILE,  # Default filename from config
                parent=self
            )
            if not report_file_path:  # User cancelled the save dialog
                messagebox.showinfo("Info", "Report generation cancelled.", parent=self)
                return

            # Write the DataFrame to an Excel file using XlsxWriter engine for formatting.
            with pd.ExcelWriter(report_file_path, engine='xlsxwriter') as writer:
                open_invoices_df.to_excel(writer, index=False, sheet_name='Open Invoices')
                workbook = writer.book
                worksheet = writer.sheets['Open Invoices']
                
                # Apply currency formatting to specified columns.
                currency_format_excel = workbook.add_format({'num_format': '$#,##0.00'})
                for col_name in config.CURRENCY_COLUMNS:
                    if col_name in open_invoices_df.columns:
                        col_idx = open_invoices_df.columns.get_loc(col_name)
                        worksheet.set_column(col_idx, col_idx, None, currency_format_excel)
                
                # Auto-adjust column widths based on content.
                for i, col in enumerate(open_invoices_df.columns):
                    column_len = max(open_invoices_df[col].astype(str).map(len).max(skipna=True), len(col)) + 2
                    worksheet.set_column(i, i, column_len if pd.notna(column_len) else len(col) + 2)
            
            messagebox.showinfo("Success", f"Report generated: {report_file_path}", parent=self)
        except Exception as e:
            logging.error(f"AppShell: Error generating report: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred generating the report: {e}", parent=self)

    def perform_data_update(self, df_row_index, column_name, new_value):
        """
        Updates a specific cell in the main status_df.
        This method is typically called by the DataManagementTab when a user edits a cell.
        Args:
            df_row_index: The DataFrame index of the row to update.
            column_name: The name of the column to update.
            new_value: The new value for the cell.
        Raises:
            IndexError: If the df_row_index is invalid.
            Exception: Re-raises other exceptions for the caller to handle.
        """
        try:
            if self.status_df is not None and df_row_index in self.status_df.index:
                self.status_df.loc[df_row_index, column_name] = new_value
                logging.info(f"AppShell: Data updated for index {df_row_index}, column '{column_name}'.")
                # The DataManagementTab is responsible for calling self.app.notify_data_changed()
                # after its UI (Treeview) has been updated.
            else:
                logging.error(f"AppShell: Invalid index {df_row_index} or DataFrame not loaded for update.")
                raise IndexError(f"Invalid DataFrame index: {df_row_index}")
        except Exception as e:
            logging.error(f"AppShell: Error performing data update: {e}", exc_info=True)
            raise  # Re-raise the exception for the calling tab to handle (e.g., show an error message)

    def perform_delete_rows(self, df_indices_to_delete):
        """
        Deletes specified rows from the main status_df.
        This method is typically called by the DataManagementTab.
        Args:
            df_indices_to_delete (list): A list of DataFrame indices to delete.
        """
        if self.status_df is not None and df_indices_to_delete:
            # Filter out any indices that are not actually in the DataFrame to avoid errors.
            valid_indices = [idx for idx in df_indices_to_delete if idx in self.status_df.index]
            if not valid_indices:
                logging.warning("AppShell: No valid indices found for deletion.")
                return

            self.status_df.drop(index=valid_indices, inplace=True)
            self.status_df.reset_index(drop=True, inplace=True)  # Reset index after dropping rows
            logging.info(f"AppShell: Deleted rows with original indices: {valid_indices}")
            # The DataManagementTab is responsible for calling self.app.notify_data_changed()
            # after its UI (Treeview) has been updated.
        else:
            logging.warning("AppShell: No data to delete or DataFrame not loaded.")

    # --- Utility Methods ---
    def center_toplevel(self, toplevel_window):
        """
        Centers a Toplevel window (e.g., a popup dialog) relative
        to the main application window.
        Args:
            toplevel_window (tk.Toplevel): The Toplevel window to center.
        """
        toplevel_window.update_idletasks()  # Ensure window dimensions are up-to-date
        
        # Get dimensions and position of the main window
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()
        
        # Get dimensions of the Toplevel window
        pop_width = toplevel_window.winfo_width()
        pop_height = toplevel_window.winfo_height()
        
        # Calculate coordinates for centering
        x = main_x + (main_width // 2) - (pop_width // 2)
        y = main_y + (main_height // 2) - (pop_height // 2)
        
        toplevel_window.geometry(f"+{x}+{y}")  # Set the Toplevel window's position

    def quit_app(self):
        """
        Handles the application quit process.
        Prompts the user to save changes before exiting.
        """
        # Ask the user if they want to save changes, cancel, or quit without saving.
        user_choice = messagebox.askyesnocancel(
            "Quit",
            "Do you want to save changes before quitting?",
            parent=self
        )
        
        if user_choice is True:  # User chose "Yes" (save and quit)
            self.save_current_data()
            self.destroy()  # Close the application
        elif user_choice is False:  # User chose "No" (quit without saving)
            self.destroy()
        # If user_choice is None (user chose "Cancel"), do nothing and keep the app running.

# --- Main Execution Block ---
# This block is executed if the script is run directly (e.g., python app_shell.py).
# The primary entry point for the application is usually Main.py.
if __name__ == '__main__':
    app = OpenJobsApp()  # Create an instance of the application
    app.mainloop()      # Start the Tkinter event loop
