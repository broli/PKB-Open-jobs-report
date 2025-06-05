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
from export_tab import ExportTab 

# --- Logging Configuration ---
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
        super().__init__()

        sv_ttk.set_theme("light")
        # Window title now reflects APP_VERSION from config (e.g., "2.0.0")
        self.title(config.APP_NAME + " - v" + config.APP_VERSION) 
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

        self.DEFAULT_FONT = config.DEFAULT_FONT
        self.DEFAULT_FONT_BOLD = config.DEFAULT_FONT_BOLD
        self.CURRENCY_FORMAT = config.CURRENCY_FORMAT
        
        default_font_string = f"{config.DEFAULT_FONT_FAMILY} {config.DEFAULT_FONT_SIZE}"
        self.option_add("*Font", default_font_string)
        self.option_add("*Text*Font", default_font_string)
        self.option_add("*Label*Font", default_font_string)
        self.option_add("*Button*Font", default_font_string)
        self.option_add("*Menu*Font", default_font_string)
        self.option_add("*MenuItem*Font", default_font_string)
        self.option_add("*TCombobox*Listbox*Font", default_font_string)

        self.maximize_window()

        self.style = ttk.Style(self)
        self.style.configure('.', font=config.DEFAULT_FONT)
        self.style.map("Treeview",
                       background=[('selected', config.STATUS_COLORS["selected_bg"])],
                       foreground=[('selected', config.STATUS_COLORS["selected_fg"])])

        self.status_df = None
        self.load_initial_data() # Calls updated data_utils.load_status() for SQLite

        self.notebook = None
        self.data_tab_instance = None
        self.reporting_tab_instance = None
        self.export_tab_instance = None
        
        self.create_main_ui_layout()
        
        if self.data_tab_instance and self.status_df is not None:
            self.data_tab_instance.populate_treeview()

    def load_initial_data(self):
        """
        Loads initial status data. Now uses data_utils.load_status() which reads from SQLite.
        """
        self.status_df = data_utils.load_status() # This now loads from SQLite
        
        if self.status_df is None: # Should be an empty DataFrame if DB/table not found
            messagebox.showerror("Fatal Error", "Could not load or initialize status data. Exiting.")
            self.destroy()
            return
            
        missing_cols = [col for col in config.EXPECTED_COLUMNS if col not in self.status_df.columns]
        if missing_cols:
            logging.warning(f"AppShell: Loaded data is missing columns: {', '.join(missing_cols)}. Adding them.")
            for col in missing_cols:
                # Ensure date columns are initialized as NaT if missing, others as None
                if col in ['Order Date', 'Turn in Date']:
                    self.status_df[col] = pd.NaT
                else:
                    self.status_df[col] = None
        
        self.status_df = self.status_df.reindex(columns=config.EXPECTED_COLUMNS)
        # Ensure date columns are of datetime type after reindexing and potential additions
        for col_name in ['Order Date', 'Turn in Date']:
            if col_name in self.status_df.columns:
                self.status_df[col_name] = pd.to_datetime(self.status_df[col_name], errors='coerce')


    def maximize_window(self):
        try:
            self.state('zoomed')
        except tk.TclError:
            try:
                m = self.maxsize()
                self.geometry('{}x{}+0+0'.format(*m))
            except tk.TclError:
                self.attributes('-fullscreen', True)

    def create_main_ui_layout(self):
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel_data)
        file_menu.add_command(label="Save Current Status", command=self.save_current_data) # Saves to SQLite
        # REMOVED: file_menu.add_command(label="Generate Open Jobs Report (Excel)", command=self.generate_excel_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit_app)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # NEW: Help Menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about_dialog)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menu_bar)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=config.DEFAULT_PADDING, pady=(0, config.DEFAULT_PADDING))

        self.data_tab_instance = DataManagementTab(self.notebook, self)
        self.notebook.add(self.data_tab_instance, text='Data Management')

        self.reporting_tab_instance = ReportingTab(self.notebook, self)
        self.notebook.add(self.reporting_tab_instance, text='Reporting & Statistics')
        
        self.export_tab_instance = ExportTab(self.notebook, self)
        self.notebook.add(self.export_tab_instance, text='Export Report')
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_notebook_tab_changed)

    # NEW: Method to display the About dialog
    def show_about_dialog(self):
        """Displays an 'About' dialog with application information."""
        about_message = (
            f"{config.APP_NAME}\n"
            f"Version: {config.APP_VERSION}\n\n"
            "This application helps track and manage open job statuses.\n"
            "Working to also add reporting and meta data analysis\n\n"
            "Developed by: Carlos Ferrabone" # You can customize this
        )
        messagebox.showinfo("About " + config.APP_NAME, about_message, parent=self)

    def on_notebook_tab_changed(self, event=None):
        if not self.notebook:
            return
        try:
            selected_tab_widget = self.notebook.nametowidget(self.notebook.select())
            logging.debug(f"AppShell: Tab changed to: {self.notebook.tab(self.notebook.select(), 'text')}")
            
            if hasattr(selected_tab_widget, 'on_tab_selected'):
                selected_tab_widget.on_tab_selected()
        except tk.TclError:
            logging.warning("AppShell: Error identifying selected tab during NotebookTabChanged event.")

    def notify_data_changed(self):
        logging.debug("AppShell: Data changed, notifying relevant tabs.")
        if self.reporting_tab_instance and hasattr(self.reporting_tab_instance, 'on_tab_selected'):
            self.reporting_tab_instance.on_tab_selected() 
        
        if self.export_tab_instance and hasattr(self.export_tab_instance, 'on_tab_selected'):
             self.export_tab_instance.on_tab_selected()
        # DataManagementTab usually refreshes itself via populate_treeview after its own operations.
        # If a global refresh is needed for it from an external data change, that could be added here.


    def load_new_excel_data(self):
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")],
            parent=self
        )
        if not excel_file_path:
            return

        new_df_raw = data_utils.load_excel(excel_file_path) # Handles 'Invoice #'
        if new_df_raw is None:
            return

        current_status_df_copy = self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=config.EXPECTED_COLUMNS)
        
        processed_df = None
        try:
            # process_data uses 'Invoice #' and handles data now coming from SQLite backed status_df
            processed_df = data_utils.process_data(new_df_raw, current_status_df_copy) 

            if processed_df is None:
                messagebox.showerror("Processing Error",
                                     "Failed to process the new Excel data. The data might be invalid or incomplete.",
                                     parent=self)
                logging.error("AppShell: data_utils.process_data returned None.")
                return

        except KeyError as e:
            expected_columns_display = " | ".join(config.EXPECTED_COLUMNS)
            error_message = (
                f"A column anmed '{e}' was expected but not found in the Excel file.\n\n"
                f"Please ensure your Excel file includes all required columns with exact names (case-sensitive) "
                f"in the header row, and that data starts directly under it.\n\n"
                f"Expected column names:\n"
                f"{expected_columns_display}\n\n"
                f"The operation has been cancelled."
            )
            messagebox.showerror("Excel Format Issue", error_message, parent=self)
            logging.error(f"AppShell: KeyError during Excel data processing: Column '{e}' not found or other column mismatch.", exc_info=True)
            return

        except Exception as e:
            messagebox.showerror("Processing Error",
                                 f"An unexpected error occurred while processing the Excel data: {e}",
                                 parent=self)
            logging.error(f"AppShell: Unexpected error during Excel data processing: {e}", exc_info=True)
            return
        
        self.status_df = processed_df.reindex(columns=config.EXPECTED_COLUMNS)
        # Ensure date columns are datetime after processing and reindexing
        for col_name in ['Order Date', 'Turn in Date']:
            if col_name in self.status_df.columns:
                self.status_df[col_name] = pd.to_datetime(self.status_df[col_name], errors='coerce')

        if self.data_tab_instance:
            self.data_tab_instance.populate_treeview()
        messagebox.showinfo("Success", "New Excel data loaded and processed successfully.", parent=self)
        self.notify_data_changed()

    def save_current_data(self):
        """Saves the current status_df to SQLite via data_utils.save_status()."""
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.", parent=self)
            return
        data_utils.save_status(self.status_df) # Now saves to SQLite

    # REMOVED: generate_excel_report(self) method

    def perform_data_update(self, df_row_index, column_name, new_value):
        try:
            if self.status_df is not None and df_row_index in self.status_df.index:
                # Ensure that if a date column is updated, the value is appropriately typed if possible
                if column_name in ['Order Date', 'Turn in Date']:
                    new_value = pd.to_datetime(new_value, errors='coerce') # Coerce to NaT if unparseable

                self.status_df.loc[df_row_index, column_name] = new_value
                logging.info(f"AppShell: Data updated for index {df_row_index}, column '{column_name}'.")
            else:
                logging.error(f"AppShell: Invalid index {df_row_index} or DataFrame not loaded for update.")
                raise IndexError(f"Invalid DataFrame index: {df_row_index}")
        except Exception as e:
            logging.error(f"AppShell: Error performing data update: {e}", exc_info=True)
            raise

    def perform_delete_rows(self, df_indices_to_delete):
        if self.status_df is not None and df_indices_to_delete:
            valid_indices = [idx for idx in df_indices_to_delete if idx in self.status_df.index]
            if not valid_indices:
                logging.warning("AppShell: No valid indices found for deletion.")
                return

            self.status_df.drop(index=valid_indices, inplace=True)
            self.status_df.reset_index(drop=True, inplace=True)
            logging.info(f"AppShell: Deleted rows with original indices: {valid_indices}")
        else:
            logging.warning("AppShell: No data to delete or DataFrame not loaded.")

    def center_toplevel(self, toplevel_window):
        toplevel_window.update_idletasks()
        
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()
        
        pop_width = toplevel_window.winfo_width()
        pop_height = toplevel_window.winfo_height()
        
        x = main_x + (main_width // 2) - (pop_width // 2)
        y = main_y + (main_height // 2) - (pop_height // 2)
        
        toplevel_window.geometry(f"+{x}+{y}")

    def quit_app(self):
        user_choice = messagebox.askyesnocancel(
            "Quit",
            "Do you want to save changes before quitting?",
            parent=self
        )
        
        if user_choice is True:
            self.save_current_data() # Saves to SQLite
            self.destroy()
        elif user_choice is False:
            self.destroy()

if __name__ == '__main__':
    app = OpenJobsApp()
    app.mainloop()