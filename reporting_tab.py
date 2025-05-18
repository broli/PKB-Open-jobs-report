# reporting_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import logging

import config

class ReportingTab(ttk.Frame):
    def __init__(self, parent_notebook, app_instance):
        super().__init__(parent_notebook)
        self.app = app_instance
        self.stats_notebook = None 
        self.overall_stats_text_area = None
        self.coordinator_tabs_widgets = {} 

        self._setup_ui()
        # Tags are now configured per widget in _configure_tags_for_text_widget

    def _configure_tags_for_text_widget(self, text_widget):
        """Helper to configure standard tags on a given text widget."""
        if not text_widget: return
        # Using REPORT_TEXT_FG_COLOR from config for default text, if defined, else black
        default_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black')

        text_widget.tag_configure("header", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE + 2, "bold", "underline"), foreground="#003366") # Dark Blue
        text_widget.tag_configure("subheader", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE + 1, "bold"), foreground=default_fg)
        text_widget.tag_configure("bold_metric", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE, "bold"), foreground=default_fg)
        text_widget.tag_configure("key_value_label", foreground="#404040") # Dark Gray for labels
        text_widget.tag_configure("indented_item", lmargin1="20p", lmargin2="20p", foreground=default_fg) 
        text_widget.tag_configure("warning_text", foreground="#990000", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE, "italic")) # Dark Red Italic
        # Ensure all text inserted without explicit tags uses the default_fg
        text_widget.config(fg=default_fg)


    def _insert_text_with_tags(self, text_widget, text, tags=None):
        """Helper to insert text with optional tags."""
        # Ensure a newline is added only if the text doesn't end with one, 
        # and we want a distinct line break.
        processed_text = text
        if not isinstance(tags, tuple) and tags is not None: # Ensure tags is a tuple if provided
            tags = (tags,)

        text_widget.insert(tk.END, processed_text, tags)
        text_widget.insert(tk.END, "\n") # Add a newline after each logical insert for structure

    def _setup_ui(self):
        main_report_frame = ttk.Frame(self, padding=config.DEFAULT_PADDING)
        main_report_frame.pack(expand=True, fill=tk.BOTH)

        title_label = ttk.Label(main_report_frame, text="Reporting and Statistics Dashboard", font=self.app.DEFAULT_FONT_BOLD)
        title_label.pack(pady=(0, 10))

        self.refresh_button = ttk.Button(main_report_frame, text="Refresh All Statistics", command=self.display_all_stats)
        self.refresh_button.pack(pady=5)

        self.stats_notebook = ttk.Notebook(main_report_frame)
        self.stats_notebook.pack(expand=True, fill=tk.BOTH, pady=5)

        overall_frame = ttk.Frame(self.stats_notebook, padding=(5,5))
        self.stats_notebook.add(overall_frame, text="Overall Pipeline Health")
        
        overall_text_frame = ttk.Frame(overall_frame)
        overall_text_frame.pack(expand=True, fill=tk.BOTH)

        # Apply background and foreground from config
        report_bg = getattr(config, 'REPORT_SUB_TAB_BG_COLOR', 'white') # Default to white if not in config
        report_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black') # Default to black

        self.overall_stats_text_area = tk.Text(
            overall_text_frame, 
            height=20, 
            width=80, 
            font=self.app.DEFAULT_FONT, 
            wrap=tk.WORD, 
            spacing1=2, spacing2=2, spacing3=2,
            bg=report_bg, # Set background color
            fg=report_fg   # Set default text color
        )
        overall_scrollbar = ttk.Scrollbar(overall_text_frame, orient="vertical", command=self.overall_stats_text_area.yview)
        self.overall_stats_text_area.configure(yscrollcommand=overall_scrollbar.set)
        self.overall_stats_text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        overall_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._configure_tags_for_text_widget(self.overall_stats_text_area) 

        self.overall_stats_text_area.insert(tk.END, "Click 'Refresh All Statistics' to load data.")
        self.overall_stats_text_area.config(state=tk.DISABLED)


    def _create_or_get_coordinator_tab(self, pc_name):
        pc_name_safe = str(pc_name).replace(".", "_dot_") 
        if pc_name_safe in self.coordinator_tabs_widgets and self.coordinator_tabs_widgets[pc_name_safe].winfo_exists():
            text_area = self.coordinator_tabs_widgets[pc_name_safe]
            text_area.config(state=tk.NORMAL)
            text_area.delete('1.0', tk.END)
            self._configure_tags_for_text_widget(text_area) 
            return text_area

        pc_frame = ttk.Frame(self.stats_notebook, padding=(5,5))
        self.stats_notebook.add(pc_frame, text=str(pc_name)[:20]) 

        pc_text_frame = ttk.Frame(pc_frame)
        pc_text_frame.pack(expand=True, fill=tk.BOTH)
        
        report_bg = getattr(config, 'REPORT_SUB_TAB_BG_COLOR', 'white')
        report_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black')

        text_area = tk.Text(
            pc_text_frame, 
            height=20, 
            width=80, 
            font=self.app.DEFAULT_FONT, 
            wrap=tk.WORD, 
            spacing1=2, spacing2=2, spacing3=2,
            bg=report_bg, # Set background color
            fg=report_fg   # Set default text color
        )
        pc_scrollbar = ttk.Scrollbar(pc_text_frame, orient="vertical", command=text_area.yview)
        text_area.configure(yscrollcommand=pc_scrollbar.set)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._configure_tags_for_text_widget(text_area) 
        self.coordinator_tabs_widgets[pc_name_safe] = text_area
        return text_area

    def _prepare_open_jobs_data(self):
        # This method remains the same as your provided version
        if self.app.status_df is None or self.app.status_df.empty:
            return None, pd.Timestamp.now().normalize()

        df_all_loaded = self.app.status_df.copy()
        today = pd.Timestamp.now().normalize()
        
        excluded_statuses = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
        open_jobs_df = df_all_loaded[~df_all_loaded['Status'].isin(excluded_statuses)].copy()

        if open_jobs_df.empty:
            return open_jobs_df, today 

        if 'Balance' in open_jobs_df.columns:
            open_jobs_df['Balance_numeric'] = pd.to_numeric(
                open_jobs_df['Balance'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                errors='coerce').fillna(0)
        if 'Invoice Total' in open_jobs_df.columns:
            open_jobs_df['InvoiceTotal_numeric'] = pd.to_numeric(
                open_jobs_df['Invoice Total'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                errors='coerce').fillna(0)

        turn_in_date_col = 'Turn in Date'
        if turn_in_date_col in open_jobs_df.columns:
            open_jobs_df['TurnInDate_dt'] = pd.to_datetime(open_jobs_df[turn_in_date_col], errors='coerce')
            valid_turn_in_dates_mask = open_jobs_df['TurnInDate_dt'].notna()
            open_jobs_df['JobAge_days'] = pd.NA
            if valid_turn_in_dates_mask.any():
                open_jobs_df.loc[valid_turn_in_dates_mask, 'JobAge_days'] = \
                    (today - open_jobs_df.loc[valid_turn_in_dates_mask, 'TurnInDate_dt']).dt.days
                open_jobs_df['JobAge_days'] = pd.to_numeric(open_jobs_df['JobAge_days'], errors='coerce')
            
            if open_jobs_df['JobAge_days'].notna().any():
                age_bins = [-1, 7, 21, 49, 56, float('inf')] 
                age_labels = ['Up to 1 Week (0-7d)', '1-3 Weeks (8-21d)', 
                              '3-7 Weeks (22-49d)', '7-8 Weeks (50-56d)', 'Over 8 Weeks (57+d)']
                open_jobs_df['Age_Bucket'] = pd.cut(open_jobs_df['JobAge_days'], bins=age_bins, labels=age_labels, right=True)
            else:
                open_jobs_df['Age_Bucket'] = pd.NA
        return open_jobs_df, today

    def display_all_stats(self):
        # This method remains the same as your provided version
        open_jobs_df, today = self._prepare_open_jobs_data()

        if open_jobs_df is None:
            self.overall_stats_text_area.config(state=tk.NORMAL)
            self.overall_stats_text_area.delete('1.0', tk.END)
            self._insert_text_with_tags(self.overall_stats_text_area, "No data available in the application.", ("warning_text",))
            self.overall_stats_text_area.config(state=tk.DISABLED)
            for pc_name_safe in list(self.coordinator_tabs_widgets.keys()):
                try:
                    tab_frame_to_forget = self.coordinator_tabs_widgets[pc_name_safe].master.master
                    for i, tab_id in enumerate(self.stats_notebook.tabs()):
                        if self.stats_notebook.nametowidget(tab_id) == tab_frame_to_forget:
                            self.stats_notebook.forget(i)
                            break
                except Exception as e: logging.warning(f"Error removing old tab for {pc_name_safe}: {e}")
                del self.coordinator_tabs_widgets[pc_name_safe]
            return

        self._populate_overall_pipeline_tab(open_jobs_df, today, len(self.app.status_df))

        project_coordinator_col = 'Project Coordinator'
        active_coordinators_safe = set()
        if project_coordinator_col in open_jobs_df.columns and not open_jobs_df.empty:
            unique_coordinators = sorted([pc for pc in open_jobs_df[project_coordinator_col].unique() if pd.notna(pc)])
            
            for pc_name in unique_coordinators:
                pc_name_safe = str(pc_name).replace(".", "_dot_")
                active_coordinators_safe.add(pc_name_safe)
                pc_specific_df = open_jobs_df[open_jobs_df[project_coordinator_col] == pc_name].copy()
                self._populate_coordinator_tab(pc_name, pc_specific_df, today) 
        
        coordinators_to_remove = set(self.coordinator_tabs_widgets.keys()) - active_coordinators_safe
        for pc_name_safe in coordinators_to_remove:
            if pc_name_safe in self.coordinator_tabs_widgets:
                try:
                    tab_frame_to_forget = self.coordinator_tabs_widgets[pc_name_safe].master.master
                    for i, tab_id in enumerate(self.stats_notebook.tabs()):
                        if self.stats_notebook.nametowidget(tab_id) == tab_frame_to_forget:
                            self.stats_notebook.forget(i)
                            break
                except Exception as e: logging.warning(f"Error removing old tab for {pc_name_safe}: {e}")
                del self.coordinator_tabs_widgets[pc_name_safe]

    def _populate_overall_pipeline_tab(self, open_jobs_df, today, num_total_jobs_loaded):
        # This method remains the same as your provided version
        txt = self.overall_stats_text_area
        txt.config(state=tk.NORMAL)
        txt.delete('1.0', tk.END)
        
        num_open_jobs = len(open_jobs_df)

        self._insert_text_with_tags(txt, f"Overall Snapshot ({today.strftime('%Y-%m-%d %H:%M:%S')})", ("header",))
        self._insert_text_with_tags(txt, "--------------------------------------------------")
        txt.insert(tk.END, "Total Jobs in Current Dataset: ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_total_jobs_loaded}", ("bold_metric",))
        txt.insert(tk.END, "Currently Open Jobs (Active): ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_open_jobs}", ("bold_metric",))
        self._insert_text_with_tags(txt, "") 

        self._insert_text_with_tags(txt, "Open Job Status Counts:", ("subheader",))
        self._insert_text_with_tags(txt, "-------------------------")
        if not open_jobs_df.empty:
            open_status_counts = open_jobs_df['Status'].value_counts()
            if not open_status_counts.empty:
                for status, count in open_status_counts.items():
                    txt.insert(tk.END, f"- {status}: ", ("indented_item", "key_value_label"))
                    self._insert_text_with_tags(txt, f"{count}", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "No open jobs with status information.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "No open jobs found.", ("indented_item",))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Financial Summary (All Open Jobs):", ("subheader",))
        self._insert_text_with_tags(txt, "------------------------------------")
        if not open_jobs_df.empty and 'Balance_numeric' in open_jobs_df.columns and 'InvoiceTotal_numeric' in open_jobs_df.columns:
            total_invoice = open_jobs_df['InvoiceTotal_numeric'].sum()
            total_balance = open_jobs_df['Balance_numeric'].sum()
            total_collected = total_invoice - total_balance
            txt.insert(tk.END, "Total Invoice Amount: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_invoice)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Collected: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_collected)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Remaining Balance: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_balance)}", ("indented_item", "bold_metric"))
        elif open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for financial summary.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Numeric financial columns not pre-calculated.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
        
        self._insert_text_with_tags(txt, "Work-in-Progress Timing (All Open Jobs, from Turn-in Date):", ("subheader",))
        self._insert_text_with_tags(txt, "-----------------------------------------------------------")
        if not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns:
            if open_jobs_df['JobAge_days'].notna().any():
                avg_age = open_jobs_df['JobAge_days'].mean()
                txt.insert(tk.END, "Average Age (since turn-in): ", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"{avg_age:.2f} days", ("indented_item", "bold_metric"))
                
                oldest_idx = open_jobs_df['JobAge_days'].idxmax()
                oldest = open_jobs_df.loc[oldest_idx]
                self._insert_text_with_tags(txt, "Oldest Project (since turn-in):", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"  - Invoice #: {oldest.get('Invoice #', 'N/A')}", ("indented_item", "bold_metric"))
                self._insert_text_with_tags(txt, f"  - Turn-in Date: {oldest['TurnInDate_dt'].strftime(config.DATE_FORMAT) if pd.notna(oldest['TurnInDate_dt']) else 'N/A'}", ("indented_item",))
                self._insert_text_with_tags(txt, f"  - Age: {oldest['JobAge_days']:.0f} days", ("indented_item", "bold_metric"))
                
                self._insert_text_with_tags(txt, "Open Job Age Distribution (since turn-in):", ("indented_item", "key_value_label"))
                if 'Age_Bucket' in open_jobs_df.columns and open_jobs_df['Age_Bucket'].notna().any():
                    bucket_counts = open_jobs_df['Age_Bucket'].value_counts().sort_index()
                    for bucket, count in bucket_counts.items():
                        txt.insert(tk.END, f"  - {bucket}: ", ("indented_item",))
                        self._insert_text_with_tags(txt, f"{count} jobs", ("indented_item", "bold_metric"))
                else: self._insert_text_with_tags(txt, "  Could not determine age distribution.", ("indented_item", "warning_text"))
            else: self._insert_text_with_tags(txt, "No valid job ages for timing summary.", ("indented_item", "warning_text"))
        elif open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for timing statistics.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Timing columns not pre-calculated.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "\"Stuck\" Jobs in Early Stages (Open Jobs > 3 Weeks in early status):", ("subheader",))
        self._insert_text_with_tags(txt, "---------------------------------------------------------------------")
        if not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns and 'Status' in open_jobs_df.columns:
            early_statuses = ["New", "Waiting Measure", "Ready to order"]
            stuck_threshold_days = 21
            stuck_jobs_df = open_jobs_df[(open_jobs_df['Status'].isin(early_statuses)) & (open_jobs_df['JobAge_days'] > stuck_threshold_days)]
            if not stuck_jobs_df.empty:
                self._insert_text_with_tags(txt, f"Found {len(stuck_jobs_df)} potentially stuck job(s):", ("indented_item", "warning_text"))
                for _, job in stuck_jobs_df.iterrows():
                    self._insert_text_with_tags(txt, f"  - Inv #: {job.get('Invoice #', 'N/A')}, Status: {job.get('Status', 'N/A')}, Age: {job.get('JobAge_days', 0):.0f}d, PC: {job.get('Project Coordinator', 'N/A')}", ("indented_item",))
            else: self._insert_text_with_tags(txt, "No open jobs identified as \"stuck\" in early stages.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine stuck jobs (missing required data).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "High-Value Aging Jobs (Open Jobs > 8 Weeks & Balance > $1000):", ("subheader",))
        self._insert_text_with_tags(txt, "----------------------------------------------------------------")
        if not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns and 'Balance_numeric' in open_jobs_df.columns:
            aging_threshold_days = 56
            value_threshold = 1000 
            hv_aging_df = open_jobs_df[(open_jobs_df['JobAge_days'] > aging_threshold_days) & (open_jobs_df['Balance_numeric'] > value_threshold)]
            if not hv_aging_df.empty:
                self._insert_text_with_tags(txt, f"Found {len(hv_aging_df)} high-value aging job(s):", ("indented_item", "warning_text"))
                for _, job in hv_aging_df.iterrows():
                    self._insert_text_with_tags(txt, f"  - Inv #: {job.get('Invoice #', 'N/A')}, Balance: {self.app.CURRENCY_FORMAT.format(job.get('Balance_numeric',0))}, Age: {job.get('JobAge_days',0):.0f}d, PC: {job.get('Project Coordinator', 'N/A')}", ("indented_item",))
            else: self._insert_text_with_tags(txt, "No high-value aging jobs identified.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine high-value aging jobs (missing required data).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Total Financial Value by Age Bucket (Open Jobs):", ("subheader",))
        self._insert_text_with_tags(txt, "------------------------------------------------")
        if not open_jobs_df.empty and 'Age_Bucket' in open_jobs_df.columns and \
           'InvoiceTotal_numeric' in open_jobs_df.columns and 'Balance_numeric' in open_jobs_df.columns and \
           open_jobs_df['Age_Bucket'].notna().any() : 
            financial_by_bucket = open_jobs_df.groupby('Age_Bucket', observed=False).agg( 
                Total_Invoice_Value=('InvoiceTotal_numeric', 'sum'),
                Total_Remaining_Balance=('Balance_numeric', 'sum'),
                Job_Count=('Invoice #', 'count') 
            )
            if not financial_by_bucket.empty:
                for bucket, data in financial_by_bucket.iterrows():
                    txt.insert(tk.END, f"- {bucket} ({data['Job_Count']} jobs):\n", ("indented_item", "key_value_label"))
                    txt.insert(tk.END, f"  - Total Invoice Value: ", ("indented_item",))
                    self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(data['Total_Invoice_Value'])}", ("indented_item", "bold_metric"))
                    txt.insert(tk.END, f"  - Total Remaining Balance: ", ("indented_item",))
                    self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(data['Total_Remaining_Balance'])}", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "No data to aggregate financial value by age bucket.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine financial value by age bucket (missing data).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        txt.config(state=tk.DISABLED)


    def _populate_coordinator_tab(self, pc_name_display, pc_open_jobs_df, today):
        pc_name_safe = str(pc_name_display).replace(".", "_dot_")
        txt = self._create_or_get_coordinator_tab(pc_name_safe) 
        
        self._insert_text_with_tags(txt, f"Statistics for {pc_name_display} ({today.strftime('%Y-%m-%d %H:%M:%S')})", ("header",))
        self._insert_text_with_tags(txt, "----------------------------------------------------------")
        txt.insert(tk.END, "Total Open Jobs Assigned: ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{len(pc_open_jobs_df)}", ("bold_metric",))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Open Jobs by Current Status:", ("subheader",))
        if not pc_open_jobs_df.empty:
            status_counts_pc = pc_open_jobs_df['Status'].value_counts()
            if not status_counts_pc.empty:
                for status, count in status_counts_pc.items():
                    txt.insert(tk.END, f"  - {status}: ", ("indented_item", "key_value_label"))
                    self._insert_text_with_tags(txt, f"{count}", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "  No open jobs with status information.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "  No open jobs assigned.", ("indented_item",))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Financial Summary (for these open jobs):", ("subheader",))
        if not pc_open_jobs_df.empty and 'Balance_numeric' in pc_open_jobs_df.columns and 'InvoiceTotal_numeric' in pc_open_jobs_df.columns:
            pc_total_invoice = pc_open_jobs_df['InvoiceTotal_numeric'].sum()
            pc_total_balance = pc_open_jobs_df['Balance_numeric'].sum()
            pc_total_collected = pc_total_invoice - pc_total_balance
            txt.insert(tk.END, "  Total Invoice Amount: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_invoice)}",("indented_item", "bold_metric"))
            txt.insert(tk.END, "  Total Collected: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_collected)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "  Total Remaining Balance: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_balance)}", ("indented_item", "bold_metric"))
        elif pc_open_jobs_df.empty: self._insert_text_with_tags(txt, "  No open jobs for financial summary.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "  Financial columns not available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Work-in-Progress Timing (for these open jobs, from Turn-in Date):", ("subheader",))
        if not pc_open_jobs_df.empty and 'JobAge_days' in pc_open_jobs_df.columns:
            if 'Age_Bucket' in pc_open_jobs_df.columns and pc_open_jobs_df['Age_Bucket'].notna().any():
                self._insert_text_with_tags(txt, "Job Age Distribution:", ("indented_item", "key_value_label"))
                pc_bucket_counts = pc_open_jobs_df['Age_Bucket'].value_counts().sort_index()
                if not pc_bucket_counts.empty:
                    for bucket, count in pc_bucket_counts.items():
                        txt.insert(tk.END, f"    - {bucket}: ", ("indented_item",)) 
                        self._insert_text_with_tags(txt, f"{count} jobs", ("indented_item", "bold_metric"))
                else: self._insert_text_with_tags(txt, "    Could not determine age distribution.", ("indented_item", "warning_text"))
            else: self._insert_text_with_tags(txt, "  No valid job ages to distribute into buckets.", ("indented_item", "warning_text"))
            self._insert_text_with_tags(txt, "")

            if pc_open_jobs_df['JobAge_days'].notna().any():
                pc_total_job_days = pc_open_jobs_df['JobAge_days'].sum()
                txt.insert(tk.END, "Total Days in Progress (sum of job ages): ", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"{pc_total_job_days:.0f} days", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "Total Days in Progress: N/A (No valid job ages)", ("indented_item", "warning_text"))
            self._insert_text_with_tags(txt, "")
        elif pc_open_jobs_df.empty: self._insert_text_with_tags(txt, "  No open jobs for timing statistics.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "  Timing data ('JobAge_days') not available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
            
        txt.config(state=tk.DISABLED)

    def on_tab_selected(self):
        logging.info("Reporting tab selected.")
        if self.app.status_df is None or self.app.status_df.empty:
             if self.overall_stats_text_area:
                self.overall_stats_text_area.config(state=tk.NORMAL)
                self.overall_stats_text_area.delete('1.0', tk.END)
                self._insert_text_with_tags(self.overall_stats_text_area, "No data loaded in the 'Data Management' tab. Please load data and click 'Refresh All Statistics'.", ("warning_text",))
                self.overall_stats_text_area.config(state=tk.DISABLED)
