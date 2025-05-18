# app_shell.py (or openJobs_class.py if you prefer to keep the name)
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sv_ttk
import pandas as pd
import logging

# Import configurations and utility functions
import config
import data_utils # data_utils uses config now

# Import Tab classes
from data_management_tab import DataManagementTab
from reporting_tab import ReportingTab

# Configure logging (this should be done once, early)
logging.basicConfig(level=config.LOG_LEVEL, format=config.LOG_FORMAT)

class OpenJobsApp(tk.Tk): # Or rename to AppShell if you like
    """
    Main application shell for the Open Jobs Status Tracker.
    Manages the main window, overall data, and orchestrates different tabs.
    """
    def __init__(self):
        super().__init__()
        sv_ttk.set_theme("light")
        self.title(config.APP_NAME + " - v" + config.APP_VERSION)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

        # --- Make common config values available to tabs via self.app ---
        self.DEFAULT_FONT = config.DEFAULT_FONT
        self.DEFAULT_FONT_BOLD = config.DEFAULT_FONT_BOLD
        self.CURRENCY_FORMAT = config.CURRENCY_FORMAT
        # EXPECTED_COLUMNS is primarily used by config, data_utils, and data_management_tab directly.
        # Tabs can import config if they need specific constants.

        default_font_string = f"{config.DEFAULT_FONT_FAMILY} {config.DEFAULT_FONT_SIZE}"
        self.option_add("*Font", default_font_string)
        # Add other global font options as before...
        self.option_add("*Text*Font", default_font_string)
        self.option_add("*Label*Font", default_font_string)
        self.option_add("*Button*Font", default_font_string)
        self.option_add("*Menu*Font", default_font_string)
        self.option_add("*MenuItem*Font", default_font_string)
        self.option_add("*TCombobox*Listbox*Font", default_font_string)


        self.maximize_window()

        self.style = ttk.Style(self)
        self.style.configure('.', font=config.DEFAULT_FONT) # Global style
        # Treeview specific styles are now primarily handled by DataManagementTab using config.STATUS_COLORS
        # However, the selected color mapping can be set globally here if it applies to all treeviews.
        self.style.map("Treeview",
                       background=[('selected', config.STATUS_COLORS["selected_bg"])],
                       foreground=[('selected', config.STATUS_COLORS["selected_fg"])])


        # --- Main Application Data ---
        self.status_df = None # Initialize
        self.load_initial_data()

        # --- UI Creation ---
        self.notebook = None
        self.data_tab_instance = None
        self.reporting_tab_instance = None
        self.create_main_ui_layout()
        
        # Initial population of the data tab's treeview
        if self.data_tab_instance and self.status_df is not None:
            self.data_tab_instance.populate_treeview()

    def load_initial_data(self):
        """Loads initial status data or creates an empty DataFrame."""
        self.status_df = data_utils.load_status()
        if self.status_df is None: # Should return empty DF from load_status on error
            messagebox.showerror("Fatal Error", "Could not load or initialize status data. Exiting.")
            self.destroy()
            return
        # Ensure DataFrame has all expected columns from config.EXPECTED_COLUMNS
        missing_cols = [col for col in config.EXPECTED_COLUMNS if col not in self.status_df.columns]
        if missing_cols:
            logging.warning(f"AppShell: Loaded data is missing columns: {', '.join(missing_cols)}. Adding them.")
            for col in missing_cols:
                self.status_df[col] = None
        self.status_df = self.status_df.reindex(columns=config.EXPECTED_COLUMNS)


    def maximize_window(self):
        try: self.state('zoomed')
        except tk.TclError:
            try:
                m = self.maxsize()
                self.geometry('{}x{}+0+0'.format(*m))
            except tk.TclError: self.attributes('-fullscreen', True)

    def create_main_ui_layout(self):
        """Creates the main UI structure: menu and notebook for tabs."""
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel_data)
        file_menu.add_command(label="Save Current Status", command=self.save_current_data)
        file_menu.add_command(label="Generate Open Jobs Report (Excel)", command=self.generate_excel_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit_app)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menu_bar)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=config.DEFAULT_PADDING, pady=(0, config.DEFAULT_PADDING))

        # Create and Add Data Management Tab
        self.data_tab_instance = DataManagementTab(self.notebook, self)
        self.notebook.add(self.data_tab_instance, text='Data Management')

        # Create and Add Reporting Tab
        self.reporting_tab_instance = ReportingTab(self.notebook, self)
        self.notebook.add(self.reporting_tab_instance, text='Reporting & Statistics')
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_notebook_tab_changed)

    def on_notebook_tab_changed(self, event=None):
        """Handles actions when a notebook tab is changed."""
        if not self.notebook: return
        try:
            selected_tab_widget = self.notebook.nametowidget(self.notebook.select())
            logging.debug(f"AppShell: Tab changed to: {self.notebook.tab(self.notebook.select(), 'text')}")
            if hasattr(selected_tab_widget, 'on_tab_selected'):
                selected_tab_widget.on_tab_selected()
        except tk.TclError:
            logging.warning("AppShell: Error identifying selected tab during NotebookTabChanged event.")


    def notify_data_changed(self):
        """Notifies relevant parts of the app that data has changed."""
        logging.debug("AppShell: Data changed, notifying reporting tab.")
        if self.reporting_tab_instance and hasattr(self.reporting_tab_instance, 'on_tab_selected'):
            # If reporting tab is active, its on_tab_selected might refresh.
            # Or, it could have a more specific refresh_data method.
            # For now, calling on_tab_selected is a general way to trigger its update logic.
            self.reporting_tab_instance.on_tab_selected() 
            # Consider if DataManagementTab also needs a specific notification beyond repopulating its own tree.


    # --- Data Operation Methods ---
    def load_new_excel_data(self):
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")],
            parent=self
        )
        if not excel_file_path: return

        new_df_raw = data_utils.load_excel(excel_file_path)
        if new_df_raw is None: return # Error handled in data_utils

        current_status_df_copy = self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=config.EXPECTED_COLUMNS)
        processed_df = data_utils.process_data(new_df_raw, current_status_df_copy)

        if processed_df is None: # Should not happen if data_utils handles errors and returns empty df
            messagebox.showerror("Error", "Failed to process new Excel data. Reverting.", parent=self)
            # self.status_df remains current_status_df_copy (or original if copy failed)
            return 
        
        self.status_df = processed_df.reindex(columns=config.EXPECTED_COLUMNS) # Ensure schema

        if self.data_tab_instance:
            self.data_tab_instance.populate_treeview() # Tell data tab to refresh its view
        messagebox.showinfo("Success", "New Excel data loaded and processed.", parent=self)
        self.notify_data_changed()

    def save_current_data(self):
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.", parent=self)
            return
        data_utils.save_status(self.status_df) # Shows its own success/error message

    def generate_excel_report(self):
        if self.status_df is None or self.status_df.empty:
            messagebox.showinfo("Info", "No data available to generate a report.", parent=self)
            return

        # Define statuses to exclude for "Open Jobs" report
        excluded_statuses = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
        open_invoices_df = self.status_df[~self.status_df['Status'].isin(excluded_statuses)].copy()
        open_invoices_df = open_invoices_df.reindex(columns=config.EXPECTED_COLUMNS)

        if open_invoices_df.empty:
            messagebox.showinfo("Info", "No open invoices to report.", parent=self)
            return

        try:
            report_file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Open Invoices Report As",
                initialfile=config.OUTPUT_FILE,
                parent=self
            )
            if not report_file_path:
                messagebox.showinfo("Info", "Report generation cancelled.", parent=self)
                return

            with pd.ExcelWriter(report_file_path, engine='xlsxwriter') as writer:
                open_invoices_df.to_excel(writer, index=False, sheet_name='Open Invoices')
                workbook = writer.book
                worksheet = writer.sheets['Open Invoices']
                currency_format_excel = workbook.add_format({'num_format': '$#,##0.00'})
                for col_name in config.CURRENCY_COLUMNS:
                    if col_name in open_invoices_df.columns:
                        col_idx = open_invoices_df.columns.get_loc(col_name)
                        worksheet.set_column(col_idx, col_idx, None, currency_format_excel)
                for i, col in enumerate(open_invoices_df.columns): # Auto-adjust column widths
                    column_len = max(open_invoices_df[col].astype(str).map(len).max(skipna=True), len(col)) + 2
                    worksheet.set_column(i, i, column_len if pd.notna(column_len) else len(col) + 2)
            messagebox.showinfo("Success", f"Report generated: {report_file_path}", parent=self)
        except Exception as e:
            logging.error(f"AppShell: Error generating report: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred generating the report: {e}", parent=self)

    def perform_data_update(self, df_row_index, column_name, new_value):
        """Updates the main status_df. Called by DataManagementTab."""
        try:
            if self.status_df is not None and df_row_index in self.status_df.index:
                self.status_df.loc[df_row_index, column_name] = new_value
                logging.info(f"AppShell: Data updated for index {df_row_index}, column '{column_name}'.")
                # self.notify_data_changed() is called by DataManagementTab after UI update
            else:
                logging.error(f"AppShell: Invalid index {df_row_index} or DataFrame not loaded for update.")
                raise IndexError(f"Invalid DataFrame index: {df_row_index}")
        except Exception as e:
            logging.error(f"AppShell: Error performing data update: {e}", exc_info=True)
            raise # Re-raise for the tab to handle or display error

    def perform_delete_rows(self, df_indices_to_delete):
        """Deletes rows from the main status_df. Called by DataManagementTab."""
        if self.status_df is not None and df_indices_to_delete:
            valid_indices = [idx for idx in df_indices_to_delete if idx in self.status_df.index]
            if not valid_indices:
                logging.warning("AppShell: No valid indices found for deletion.")
                return

            self.status_df.drop(index=valid_indices, inplace=True)
            self.status_df.reset_index(drop=True, inplace=True) # Important after drop
            logging.info(f"AppShell: Deleted rows with original indices: {valid_indices}")
            # self.notify_data_changed() is called by DataManagementTab after UI update
        else:
            logging.warning("AppShell: No data to delete or DataFrame not loaded.")


    # --- Utility Methods ---
    def center_toplevel(self, toplevel_window):
        """Centers a Toplevel window relative to this main app window."""
        toplevel_window.update_idletasks()
        main_x, main_y = self.winfo_x(), self.winfo_y()
        main_width, main_height = self.winfo_width(), self.winfo_height()
        pop_width, pop_height = toplevel_window.winfo_width(), toplevel_window.winfo_height()
        x = main_x + (main_width // 2) - (pop_width // 2)
        y = main_y + (main_height // 2) - (pop_height // 2)
        toplevel_window.geometry(f"+{x}+{y}")

    def quit_app(self):
        user_choice = messagebox.askyesnocancel("Quit", "Do you want to save changes before quitting?", parent=self)
        if user_choice is True:  # Yes
            self.save_current_data()
            self.destroy()
        elif user_choice is False:  # No
            self.destroy()
        # If user_choice is None (Cancel), do nothing.

# This is for direct execution if needed, though Main.py is the entry point
if __name__ == '__main__':
    app = OpenJobsApp() # Or AppShell() if you renamed the class
    app.mainloop()
