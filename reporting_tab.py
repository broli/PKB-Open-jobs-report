# reporting_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import logging # For logging within the tab

# Import configurations from config.py
import config 

# You might need these later for graphs:
# from matplotlib.figure import Figure
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class ReportingTab(ttk.Frame):
    def __init__(self, parent_notebook, app_instance):
        """
        Initialize the Reporting Tab.
        Args:
            parent_notebook: The ttk.Notebook widget this tab will belong to.
            app_instance: The instance of the main OpenJobsApp.
        """
        super().__init__(parent_notebook)
        self.app = app_instance  # Store a reference to the main app

        # --- UI Elements for Reporting Tab ---
        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=config.DEFAULT_PADDING, pady=config.DEFAULT_PADDING) 

        title_label = ttk.Label(main_frame, text="Reporting and Statistics", font=self.app.DEFAULT_FONT_BOLD)
        title_label.pack(pady=(0, 10))

        self.refresh_button = ttk.Button(main_frame, text="Refresh Statistics", command=self.display_summary_stats)
        self.refresh_button.pack(pady=5)

        self.stats_text_area = tk.Text(main_frame, height=15, width=80, font=self.app.DEFAULT_FONT, wrap=tk.WORD)
        stats_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.stats_text_area.yview)
        self.stats_text_area.configure(yscrollcommand=stats_scrollbar.set)
        
        self.stats_text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
    def display_summary_stats(self):
        """Calculates and displays summary statistics from the app's data."""
        self.stats_text_area.config(state=tk.NORMAL)
        self.stats_text_area.delete('1.0', tk.END)

        if self.app.status_df is None or self.app.status_df.empty:
            self.stats_text_area.insert(tk.END, "No data available.\nPlease load data in the 'Data Management' tab.")
            self.stats_text_area.config(state=tk.DISABLED)
            return

        try:
            df = self.app.status_df.copy()
            today = pd.Timestamp.now() # Get current date for age calculations

            num_total_jobs = len(df)
            excluded_statuses_for_open_count = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
            # Create a DataFrame specifically for open jobs to be reused
            open_jobs_df = df[~df['Status'].isin(excluded_statuses_for_open_count)].copy() 
            num_open_jobs = len(open_jobs_df)

            summary = f"Overall Statistics ({today.strftime('%Y-%m-%d %H:%M:%S')}):\n"
            summary += f"--------------------------------------------------\n"
            summary += f"Total Jobs Loaded (including closed/cancelled): {num_total_jobs}\n"
            summary += f"Currently Open Jobs: {num_open_jobs}\n\n"

            summary += "Job Status Counts (All Loaded Jobs):\n"
            summary += "-------------------------------------\n"
            status_counts = df['Status'].value_counts()
            if not status_counts.empty:
                for status, count in status_counts.items():
                    summary += f"- {status}: {count}\n"
            else:
                summary += "No status data available.\n"
            summary += "\n"

            # Financial Summary for OPEN JOBS
            summary += "Financial Summary (Open Jobs Only):\n"
            summary += "------------------------------------\n"
            if 'Balance' in open_jobs_df.columns and 'Invoice Total' in open_jobs_df.columns:
                try:
                    open_jobs_df['Balance_numeric'] = pd.to_numeric(
                        open_jobs_df['Balance'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                        errors='coerce'
                    ).fillna(0)
                    open_jobs_df['InvoiceTotal_numeric'] = pd.to_numeric(
                        open_jobs_df['Invoice Total'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                        errors='coerce'
                    ).fillna(0)

                    total_invoice_amount_open = open_jobs_df['InvoiceTotal_numeric'].sum()
                    total_remaining_balance_open = open_jobs_df['Balance_numeric'].sum()
                    total_collected_open = total_invoice_amount_open - total_remaining_balance_open
                    
                    summary += f"Total Invoice Amount (Open Jobs): {self.app.CURRENCY_FORMAT.format(total_invoice_amount_open)}\n"
                    summary += f"Total Collected (Open Jobs): {self.app.CURRENCY_FORMAT.format(total_collected_open)}\n"
                    summary += f"Total Remaining Balance (Open Jobs): {self.app.CURRENCY_FORMAT.format(total_remaining_balance_open)}\n\n"

                except Exception as e:
                    logging.error(f"Error calculating financial summary for open jobs: {e}", exc_info=True)
                    summary += f"Could not calculate financial summary for open jobs: {e}\n\n"
            else:
                summary += "Required columns ('Balance', 'Invoice Total') not found for financial summary.\n\n"
            
            # Jobs per Project Coordinator (for OPEN JOBS)
            project_coordinator_col = 'Project Coordinator' 
            if project_coordinator_col in open_jobs_df.columns and not open_jobs_df[project_coordinator_col].empty:
                summary += f"Open Jobs per {project_coordinator_col} (All):\n" 
                summary += "---------------------------------------------------\n"
                coordinator_counts = open_jobs_df[project_coordinator_col].value_counts() 
                if not coordinator_counts.empty:
                    for pc, count in coordinator_counts.items():
                        summary += f"- {pc if pd.notna(pc) else 'N/A'}: {count}\n" 
                else:
                    summary += f"No {project_coordinator_col} data available for open jobs.\n"
                summary += "\n"
            else:
                summary += f"'{project_coordinator_col}' column not found or empty in open jobs. Cannot show coordinator stats.\n\n"

            # --- Timing Statistics (Open Jobs Only) ---
            summary += "Timing Statistics (Open Jobs Only):\n"
            summary += "-----------------------------------\n"
            order_date_col = 'Order Date'
            if order_date_col in open_jobs_df.columns:
                # Convert 'Order Date' to datetime, coercing errors
                open_jobs_df['OrderDate_dt'] = pd.to_datetime(open_jobs_df[order_date_col], errors='coerce')
                
                # Calculate age of projects in days
                # Only calculate for rows where OrderDate_dt is a valid date (not NaT)
                valid_dates_mask = open_jobs_df['OrderDate_dt'].notna()
                if valid_dates_mask.any(): # Proceed if there's at least one valid date
                    open_jobs_df.loc[valid_dates_mask, 'ProjectAge_days'] = \
                        (today - open_jobs_df.loc[valid_dates_mask, 'OrderDate_dt']).dt.days

                    # 1. Average Age of Open Projects
                    average_age = open_jobs_df['ProjectAge_days'].mean() # NaNs are automatically excluded by mean()
                    summary += f"Average Age of Open Projects: {average_age:.2f} days\n"

                    # 2. Oldest Open Project
                    if not open_jobs_df['ProjectAge_days'].dropna().empty:
                        oldest_project_idx = open_jobs_df['ProjectAge_days'].idxmax() # Index of the oldest project
                        oldest_project = open_jobs_df.loc[oldest_project_idx]
                        summary += "Oldest Open Project:\n"
                        summary += f"  - Invoice #: {oldest_project.get('Invoice #', 'N/A')}\n"
                        summary += f"  - Account: {oldest_project.get('Account', 'N/A')}\n"
                        summary += f"  - Order Date: {oldest_project['OrderDate_dt'].strftime(config.DATE_FORMAT) if pd.notna(oldest_project['OrderDate_dt']) else 'N/A'}\n"
                        summary += f"  - Age: {oldest_project['ProjectAge_days']:.0f} days\n"
                    else:
                        summary += "No valid project ages to determine the oldest project.\n"
                    
                    # 3. Total Open Days per Project Coordinator
                    if project_coordinator_col in open_jobs_df.columns:
                        # Ensure ProjectAge_days is numeric and fill NaN with 0 for sum, or filter out NaNs before sum
                        open_jobs_df['ProjectAge_days_filled'] = open_jobs_df['ProjectAge_days'].fillna(0)
                        coordinator_age_sum = open_jobs_df.groupby(project_coordinator_col)['ProjectAge_days_filled'].sum()
                        
                        summary += f"\nTotal Open Project Days per {project_coordinator_col}:\n"
                        if not coordinator_age_sum.empty:
                            for pc, total_days in coordinator_age_sum.sort_values(ascending=False).items():
                                summary += f"  - {pc if pd.notna(pc) else 'N/A'}: {total_days:.0f} days\n"
                        else:
                            summary += f"  No {project_coordinator_col} data with valid ages available.\n"
                    else:
                        summary += f"  '{project_coordinator_col}' column not found for age summation.\n"
                else:
                    summary += "No valid 'Order Date' data found to calculate project ages.\n"
            else:
                summary += f"'{order_date_col}' column not found. Cannot calculate timing statistics.\n"
            summary += "\n"


            self.stats_text_area.insert(tk.END, summary)
        except Exception as e:
            error_message = f"Error generating statistics: {e}"
            self.stats_text_area.insert(tk.END, error_message)
            logging.error(error_message, exc_info=True)
        finally:
            self.stats_text_area.config(state=tk.DISABLED)

    def on_tab_selected(self):
        """
        This method will be called by the main app when this tab is selected.
        """
        logging.info("Reporting tab selected. Refreshing statistics.")
        self.display_summary_stats()

