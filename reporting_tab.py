# reporting_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import logging
import os # Added for potential path operations, though savefig handles full paths

import config

# --- Matplotlib Imports ---
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt 
import matplotlib.colors as mcolors # For more color options
from matplotlib.ticker import MaxNLocator # For integer ticks on y-axis

class ReportingTab(ttk.Frame):
    def __init__(self, parent_notebook, app_instance):
        super().__init__(parent_notebook)
        self.app = app_instance
        self.stats_notebook = None 
        self.overall_stats_text_area = None
        self.coordinator_tabs_widgets = {} 

        # --- For Matplotlib Charts ---
        self.overall_status_chart_figure = None # Store the figure object explicitly
        self.overall_status_chart_canvas_widget = None 
        self.overall_status_chart_frame = None 
        
        self.overall_financial_summary_chart_figure = None # Store the figure object explicitly
        self.overall_financial_summary_chart_canvas_widget = None 
        self.overall_financial_summary_chart_frame = None 
        
        self.overall_charts_canvas = None 
        self.overall_charts_scrollable_frame = None 

        # --- For New Weekly Intake Chart ---
        self.weekly_intake_tab_content_frame = None
        self.weekly_intake_chart_frame = None
        self.weekly_intake_chart_figure = None
        self.weekly_intake_chart_canvas_widget = None

        self._setup_ui()

    def _configure_tags_for_text_widget(self, text_widget):
        """Configures text tags for consistent styling in Text widgets."""
        if not text_widget: return
        default_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black')
        text_widget.tag_configure("header", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE + 2, "bold", "underline"), foreground="#003366") 
        text_widget.tag_configure("subheader", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE + 1, "bold"), foreground=default_fg)
        text_widget.tag_configure("bold_metric", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE, "bold"), foreground=default_fg)
        text_widget.tag_configure("key_value_label", foreground="#404040") 
        text_widget.tag_configure("indented_item", lmargin1="20p", lmargin2="20p", foreground=default_fg) 
        text_widget.tag_configure("warning_text", foreground="#990000", font=(config.DEFAULT_FONT_FAMILY, config.DEFAULT_FONT_SIZE, "italic")) 
        text_widget.tag_configure("separator_line_tk") # <<< ADDED TAG for separators (no visual style needed in Tk)
        text_widget.config(fg=default_fg)

    def _insert_text_with_tags(self, text_widget, text, tags=None):
        """Helper to insert text with specified tags into a Text widget."""
        processed_text = text
        if not isinstance(tags, tuple) and tags is not None: 
            tags = (tags,)
        text_widget.insert(tk.END, processed_text, tags)
        text_widget.insert(tk.END, "\n") 

    # <<< NEW METHOD for inserting separators >>>
    def _insert_separator_line(self, text_widget):
        """Inserts a themed break/separator line marker into the text widget."""
        # This will be interpreted by the HTML exporter as an <hr>
        text_widget.insert(tk.END, "\n", ("separator_line_tk",))

    def _setup_ui(self):
        """Sets up the main UI elements for the Reporting tab."""
        main_report_frame = ttk.Frame(self, padding=config.DEFAULT_PADDING)
        main_report_frame.pack(expand=True, fill=tk.BOTH)

        title_label = ttk.Label(main_report_frame, text="Reporting and Statistics Dashboard", font=self.app.DEFAULT_FONT_BOLD)
        title_label.pack(pady=(0, 10))

        self.refresh_button = ttk.Button(main_report_frame, text="Refresh All Statistics", command=self.display_all_stats)
        self.refresh_button.pack(pady=5)

        self.stats_notebook = ttk.Notebook(main_report_frame)
        self.stats_notebook.pack(expand=True, fill=tk.BOTH, pady=5)

        report_bg = getattr(config, 'REPORT_SUB_TAB_BG_COLOR', 'white') 

        # --- Overall Pipeline Health Tab ---
        overall_tab_content_frame = ttk.Frame(self.stats_notebook, padding=(5,5))
        self.stats_notebook.add(overall_tab_content_frame, text="Overall Pipeline Health")
        
        # Container for the text area (left side)
        overall_text_container_frame = tk.Frame(overall_tab_content_frame, bg=report_bg) 
        overall_text_container_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5)) 

        report_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black') 

        self.overall_stats_text_area = tk.Text(
            overall_text_container_frame, 
            height=20, 
            width=60,  # Initial width, will expand
            font=self.app.DEFAULT_FONT, 
            wrap=tk.WORD, 
            spacing1=2, spacing2=2, spacing3=2, # Line spacing
            bg=report_bg, 
            fg=report_fg   
        )
        overall_text_scrollbar = ttk.Scrollbar(overall_text_container_frame, orient="vertical", command=self.overall_stats_text_area.yview)
        self.overall_stats_text_area.configure(yscrollcommand=overall_text_scrollbar.set)
        self.overall_stats_text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        overall_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._configure_tags_for_text_widget(self.overall_stats_text_area) 
        self.overall_stats_text_area.insert(tk.END, "Click 'Refresh All Statistics' to load data.")
        self.overall_stats_text_area.config(state=tk.DISABLED) # Read-only initially

        # Container for charts (right side)
        charts_outer_frame = ttk.Frame(overall_tab_content_frame, padding=(5,0))
        charts_outer_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Scrollable canvas for charts
        self.overall_charts_canvas = tk.Canvas(charts_outer_frame, bg=report_bg) 
        self.overall_charts_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        charts_v_scrollbar = ttk.Scrollbar(charts_outer_frame, orient="vertical", command=self.overall_charts_canvas.yview)
        charts_v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.overall_charts_canvas.configure(yscrollcommand=charts_v_scrollbar.set)
        
        # Frame inside the canvas that will hold the charts
        self.overall_charts_scrollable_frame = tk.Frame(self.overall_charts_canvas, bg=report_bg) 
        self.overall_charts_canvas.create_window((0, 0), window=self.overall_charts_scrollable_frame, anchor="nw", tags="scrollable_frame_tag")

        # Update scrollregion when the scrollable frame's size changes
        self.overall_charts_scrollable_frame.bind("<Configure>", lambda e: self.overall_charts_canvas.configure(scrollregion=self.overall_charts_canvas.bbox("all")))

        # Placeholder frames for charts (actual charts will be drawn here)
        self.overall_status_chart_frame = ttk.Frame(self.overall_charts_scrollable_frame, padding=(5,5))
        self.overall_status_chart_frame.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(0,10)) 
        ttk.Label(self.overall_status_chart_frame, text="Open Jobs by Status (Chart will appear here)").pack() 

        self.overall_financial_summary_chart_frame = ttk.Frame(self.overall_charts_scrollable_frame, padding=(5,5))
        self.overall_financial_summary_chart_frame.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(0,10)) 
        ttk.Label(self.overall_financial_summary_chart_frame, text="Financial Summary (Chart will appear here)").pack() 

        # --- Weekly Job Intake Tab (New) ---
        self.weekly_intake_tab_content_frame = ttk.Frame(self.stats_notebook, padding=(5,5))
        self.stats_notebook.add(self.weekly_intake_tab_content_frame, text="Weekly Job Intake")

        # Frame for the chart itself (will be filled by _create_weekly_intake_chart)
        self.weekly_intake_chart_frame = ttk.Frame(self.weekly_intake_tab_content_frame, padding=(5,5))
        self.weekly_intake_chart_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0,10))
        ttk.Label(self.weekly_intake_chart_frame, text="Weekly Job Intake Chart (Data will load upon refresh)").pack()


    def _create_or_get_coordinator_tab(self, pc_name):
        """Creates a new tab for a Project Coordinator or returns the existing Text widget if it exists."""
        pc_name_safe = str(pc_name).replace(".", "_dot_") # Sanitize name for widget path
        
        # If tab and text area already exist, clear and return it
        if pc_name_safe in self.coordinator_tabs_widgets and self.coordinator_tabs_widgets[pc_name_safe].winfo_exists():
            text_area = self.coordinator_tabs_widgets[pc_name_safe]
            text_area.config(state=tk.NORMAL)
            text_area.delete('1.0', tk.END)
            self._configure_tags_for_text_widget(text_area) # Re-apply styles
            return text_area 

        # Create new tab frame
        pc_tab_frame = ttk.Frame(self.stats_notebook, padding=(5,5)) 
        self.stats_notebook.add(pc_tab_frame, text=str(pc_name)[:20]) # Truncate long names for tab label

        # Frame to hold text area and scrollbar
        pc_text_content_frame = tk.Frame(pc_tab_frame) 
        pc_text_content_frame.pack(expand=True, fill=tk.BOTH)
        
        report_bg = getattr(config, 'REPORT_SUB_TAB_BG_COLOR', 'white')
        report_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black')

        text_area = tk.Text(
            pc_text_content_frame, 
            height=20, width=80, 
            font=self.app.DEFAULT_FONT, wrap=tk.WORD, 
            spacing1=2, spacing2=2, spacing3=2,
            bg=report_bg, fg=report_fg   
        )
        pc_scrollbar = ttk.Scrollbar(pc_text_content_frame, orient="vertical", command=text_area.yview)
        text_area.configure(yscrollcommand=pc_scrollbar.set)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._configure_tags_for_text_widget(text_area) 
        self.coordinator_tabs_widgets[pc_name_safe] = text_area # Store reference
        return text_area 

    def _prepare_open_jobs_data(self):
        """
        Prepares the 'open jobs' DataFrame for reporting.
        Filters out closed/cancelled jobs, calculates numeric financial columns,
        and derives job age and age buckets.
        'Balance_numeric' here represents the OUTSTANDING balance from the source data.
        Returns:
            pd.DataFrame: The processed open jobs DataFrame.
            pd.Timestamp: The current timestamp (normalized to day).
        """
        if self.app.status_df is None or self.app.status_df.empty:
            logging.warning("ReportingTab: No status_df available for processing.")
            return None, pd.Timestamp.now().normalize()

        df_all_loaded = self.app.status_df.copy()
        # Ensure date columns are datetime before calculations
        date_cols_to_convert = ['Turn in Date', 'Order Date'] # Add other relevant date cols if needed
        for col in date_cols_to_convert:
            if col in df_all_loaded.columns:
                df_all_loaded[col] = pd.to_datetime(df_all_loaded[col], errors='coerce')


        today = pd.Timestamp.now().normalize() # For consistent age calculation

        # Define statuses to exclude for "open jobs"
        excluded_statuses = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
        open_jobs_df = df_all_loaded[~df_all_loaded['Status'].isin(excluded_statuses)].copy()

        if open_jobs_df.empty:
            logging.info("ReportingTab: No open jobs after filtering.")
            return open_jobs_df, today 

        # Convert currency columns to numeric, handling errors
        # 'Balance' column is OUTSTANDING balance
        if 'Balance' in open_jobs_df.columns:
            open_jobs_df['Balance_numeric'] = pd.to_numeric(
                open_jobs_df['Balance'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                errors='coerce').fillna(0)
        if 'Invoice Total' in open_jobs_df.columns:
            open_jobs_df['InvoiceTotal_numeric'] = pd.to_numeric(
                open_jobs_df['Invoice Total'].astype(str).replace({'\$': '', ',': ''}, regex=True),
                errors='coerce').fillna(0)

        # Calculate Job Age and Age Buckets
        turn_in_date_col = 'Turn in Date'
        if turn_in_date_col in open_jobs_df.columns:
            open_jobs_df['TurnInDate_dt'] = pd.to_datetime(open_jobs_df[turn_in_date_col], errors='coerce')
            
            valid_turn_in_dates_mask = open_jobs_df['TurnInDate_dt'].notna()
            open_jobs_df['JobAge_days'] = pd.NA # Initialize column

            if valid_turn_in_dates_mask.any():
                open_jobs_df.loc[valid_turn_in_dates_mask, 'JobAge_days'] = \
                    (today - open_jobs_df.loc[valid_turn_in_dates_mask, 'TurnInDate_dt']).dt.days
                open_jobs_df['JobAge_days'] = pd.to_numeric(open_jobs_df['JobAge_days'], errors='coerce') # Ensure numeric

            if open_jobs_df['JobAge_days'].notna().any():
                age_bins = [-1, 7, 21, 49, 56, float('inf')] 
                age_labels = [
                    'Job Age: 0-7 Days (First Week of Job Age)', 
                    'Job Age: 8-21 Days (Job is 2-3 Weeks Old)', 
                    'Job Age: 22-49 Days (Job is 4-7 Weeks Old)', 
                    'Job Age: 50-56 Days (Job is 8 Weeks Old)', 
                    'Job Age: Over 56 Days (Job is Older than 8 Weeks)'
                ]
                open_jobs_df['Age_Bucket'] = pd.cut(open_jobs_df['JobAge_days'], bins=age_bins, labels=age_labels, right=True)
            else:
                open_jobs_df['Age_Bucket'] = pd.NA 
        else:
            logging.warning("ReportingTab: 'Turn in Date' column not found. Cannot calculate job age.")
            open_jobs_df['JobAge_days'] = pd.NA
            open_jobs_df['Age_Bucket'] = pd.NA
            
        return open_jobs_df, today

    def _prepare_weekly_intake_data(self, source_df):
        """
        Prepares data for the weekly job intake chart, focusing on the last 3 months.
        Args:
            source_df (pd.DataFrame): The DataFrame to use (e.g., self.app.status_df).
        Returns:
            pd.Series | None: A Series with weekly job counts, or None if data is unavailable.
        """
        if source_df is None or source_df.empty:
            logging.warning("ReportingTab: Source DataFrame for weekly intake is empty or None.")
            return None
        
        if 'Turn in Date' not in source_df.columns:
            logging.warning("ReportingTab: 'Turn in Date' column missing, cannot generate weekly intake data.")
            return None

        df_copy = source_df.copy()
        df_copy['TurnInDate_dt'] = pd.to_datetime(df_copy['Turn in Date'], errors='coerce')
        df_copy = df_copy.dropna(subset=['TurnInDate_dt'])

        if df_copy.empty:
            logging.info("ReportingTab: No valid 'Turn in Date' entries after conversion for weekly intake.")
            return None

        three_months_ago = pd.Timestamp.now().normalize() - pd.DateOffset(months=3)
        df_copy = df_copy[df_copy['TurnInDate_dt'] >= three_months_ago]
        
        if df_copy.empty:
            logging.info("ReportingTab: No 'Turn in Date' entries in the last 3 months for weekly intake.")
            return None
            
        weekly_counts = df_copy.set_index('TurnInDate_dt').resample('W-MON').size()
        weekly_counts = weekly_counts.rename("Job Count")
        
        logging.info(f"ReportingTab: Prepared weekly intake data for the last 3 months with {len(weekly_counts)} weeks.")
        return weekly_counts

    def _create_weekly_intake_chart(self, parent_frame, intake_data):
        """
        Creates and embeds a bar chart for weekly job intake.
        Args:
            parent_frame (tk.Frame): The Tkinter frame to embed the chart in.
            intake_data (pd.Series): Data prepared by _prepare_weekly_intake_data.
        """
        if hasattr(self, 'weekly_intake_chart_canvas_widget') and self.weekly_intake_chart_canvas_widget:
            self.weekly_intake_chart_canvas_widget.get_tk_widget().destroy()
            self.weekly_intake_chart_canvas_widget = None
            self.weekly_intake_chart_figure = None 
        for widget in parent_frame.winfo_children(): 
            widget.destroy()

        if intake_data is None or intake_data.empty:
            ttk.Label(parent_frame, text="No data available to display for weekly job intake (last 3 months).").pack(expand=True, fill=tk.BOTH)
            logging.info("ReportingTab: No intake data to plot for weekly chart (last 3 months).")
            return

        try:
            num_weeks = len(intake_data)
            fig_width = max(8, num_weeks * 0.5) 
            fig_height = 5 

            self.weekly_intake_chart_figure = Figure(figsize=(fig_width, fig_height), dpi=90)
            ax = self.weekly_intake_chart_figure.add_subplot(111)
            
            intake_data.plot(kind='bar', ax=ax, color=plt.cm.get_cmap('viridis', num_weeks)(range(num_weeks)))
            
            ax.set_title('Weekly Job Intake (Last 3 Months by Turn in Date)', fontsize=config.DEFAULT_FONT_SIZE + 1)
            ax.set_xlabel('Week of Year (Monday Start)', fontsize=config.DEFAULT_FONT_SIZE)
            ax.set_ylabel('Number of Jobs Received', fontsize=config.DEFAULT_FONT_SIZE)
            
            ax.set_xticklabels([f"Week {date.strftime('%W')}" for date in intake_data.index], rotation=45, ha="right")
            
            ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.grid(axis='y', linestyle=':', alpha=0.7, color='gray')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            self.weekly_intake_chart_figure.tight_layout()

            self.weekly_intake_chart_canvas_widget = FigureCanvasTkAgg(self.weekly_intake_chart_figure, master=parent_frame)
            self.weekly_intake_chart_canvas_widget.draw()
            canvas_tk_widget = self.weekly_intake_chart_canvas_widget.get_tk_widget()
            canvas_tk_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            
            logging.info(f"ReportingTab: Successfully created weekly intake chart with {num_weeks} weeks (last 3 months).")

        except Exception as e:
            logging.error(f"ReportingTab: Error creating weekly intake chart: {e}", exc_info=True)
            ttk.Label(parent_frame, text=f"Error creating weekly intake chart: {e}").pack(expand=True, fill=tk.BOTH)
            self.weekly_intake_chart_figure = None


    def _create_status_distribution_chart(self, parent_frame, open_jobs_df):
        """Creates and embeds a horizontal bar chart for status distribution."""
        if hasattr(self, 'overall_status_chart_canvas_widget') and self.overall_status_chart_canvas_widget:
            self.overall_status_chart_canvas_widget.get_tk_widget().destroy()
            self.overall_status_chart_canvas_widget = None
            self.overall_status_chart_figure = None 
        for widget in parent_frame.winfo_children(): widget.destroy() 

        if open_jobs_df is None or open_jobs_df.empty or 'Status' not in open_jobs_df.columns:
            ttk.Label(parent_frame, text="No data for status chart.").pack(expand=True, fill=tk.BOTH)
            return
        
        status_counts = open_jobs_df['Status'].value_counts().sort_index()
        if status_counts.empty:
            ttk.Label(parent_frame, text="No status data to plot.").pack(expand=True, fill=tk.BOTH)
            return

        try:
            fig_height = max(3.5, len(status_counts) * 0.5) 
            self.overall_status_chart_figure = Figure(figsize=(6, fig_height), dpi=90) 
            ax = self.overall_status_chart_figure.add_subplot(111)
            
            colors = plt.cm.get_cmap('Pastel2', len(status_counts)) 
            bars = status_counts.plot(kind='barh', ax=ax, color=[colors(i) for i in range(len(status_counts))])
            
            ax.set_title('Open Jobs by Status', fontsize=config.DEFAULT_FONT_SIZE +1)
            ax.set_xlabel('Number of Open Jobs', fontsize=config.DEFAULT_FONT_SIZE)
            ax.set_ylabel('Status', fontsize=config.DEFAULT_FONT_SIZE)
            ax.tick_params(axis='both', which='major', labelsize=config.DEFAULT_FONT_SIZE - 1)
            
            ax.grid(axis='x', linestyle=':', alpha=0.7, color='gray') 
            ax.spines['top'].set_visible(False) 
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('darkgrey')
            ax.spines['bottom'].set_color('darkgrey')

            self.overall_status_chart_figure.subplots_adjust(left=0.30, right=0.95, top=0.90, bottom=0.15) 
            
            for i, v in enumerate(status_counts):
                ax.text(v + 0.2, i, str(v), color='black', va='center', fontweight='normal', fontsize=config.DEFAULT_FONT_SIZE -1)
            
            self.overall_status_chart_canvas_widget = FigureCanvasTkAgg(self.overall_status_chart_figure, master=parent_frame)
            self.overall_status_chart_canvas_widget.draw()
            canvas_tk_widget = self.overall_status_chart_canvas_widget.get_tk_widget()
            canvas_tk_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True) 
            parent_frame.update_idletasks() 
            
            self.overall_charts_canvas.after_idle(lambda: self.overall_charts_canvas.configure(scrollregion=self.overall_charts_canvas.bbox("all")))

        except Exception as e:
            logging.error(f"ReportingTab: Error creating status distribution chart: {e}", exc_info=True)
            ttk.Label(parent_frame, text=f"Error creating chart: {e}").pack(expand=True, fill=tk.BOTH)
            self.overall_status_chart_figure = None 

    def _create_financial_summary_chart(self, parent_frame, open_jobs_df):
        """
        Creates and embeds a pie chart for the financial summary.
        Assumes 'Balance_numeric' is OUTSTANDING balance.
        """
        if hasattr(self, 'overall_financial_summary_chart_canvas_widget') and self.overall_financial_summary_chart_canvas_widget:
            self.overall_financial_summary_chart_canvas_widget.get_tk_widget().destroy()
            self.overall_financial_summary_chart_canvas_widget = None
            self.overall_financial_summary_chart_figure = None 
        for widget in parent_frame.winfo_children(): widget.destroy()

        if open_jobs_df is None or open_jobs_df.empty or \
           'InvoiceTotal_numeric' not in open_jobs_df.columns or \
           'Balance_numeric' not in open_jobs_df.columns: # Balance_numeric is outstanding
            ttk.Label(parent_frame, text="No data for financial summary chart.").pack(expand=True, fill=tk.BOTH)
            return

        total_invoice = open_jobs_df['InvoiceTotal_numeric'].sum()
        total_outstanding_balance = open_jobs_df['Balance_numeric'].sum() # Sum of 'Balance' column
        total_collected_calculated = total_invoice - total_outstanding_balance
        
        pie_labels_for_legend = []
        pie_values = []
        pie_colors = []
        explode_values = [] 

        if total_collected_calculated > 0:
            pie_labels_for_legend.append(f'Collected (${total_collected_calculated:,.0f})')
            pie_values.append(total_collected_calculated)
            pie_colors.append('#77DD77') 
            explode_values.append(0.03) 
        if total_outstanding_balance > 0:
            pie_labels_for_legend.append(f'Remaining (${total_outstanding_balance:,.0f})')
            pie_values.append(total_outstanding_balance)
            pie_colors.append('#FFB347') 
            explode_values.append(0.03)

        if not pie_values or sum(pie_values) == 0:
             ttk.Label(parent_frame, text="Financial values are zero or not suitable for pie chart.").pack(expand=True, fill=tk.BOTH)
             return

        try:
            self.overall_financial_summary_chart_figure = Figure(figsize=(6, 4.5), dpi=90) 
            ax = self.overall_financial_summary_chart_figure.add_subplot(111) 
            
            wedges, texts, autotexts = ax.pie(
                pie_values, 
                autopct='%1.1f%%', 
                startangle=90, 
                colors=pie_colors,
                explode=explode_values,
                wedgeprops=dict(width=0.45, edgecolor='w'), 
                pctdistance=0.75, 
                textprops={'fontsize': config.DEFAULT_FONT_SIZE - 1, 'color':'black', 'weight':'bold'}
            )
            
            ax.set_title(f'Open Jobs Financials\n(Total Invoice: ${total_invoice:,.0f})', fontsize=config.DEFAULT_FONT_SIZE, pad=15) 
            ax.axis('equal')  
            
            lgd = ax.legend(wedges, pie_labels_for_legend,
                      title="Breakdown",
                      loc="center left", 
                      bbox_to_anchor=(0.78, 0.90), 
                      fontsize=config.DEFAULT_FONT_SIZE - 1,
                      title_fontsize=config.DEFAULT_FONT_SIZE,
                      facecolor='aliceblue')

            self.overall_financial_summary_chart_figure.subplots_adjust(left=0.05, bottom=0.05, right=0.70, top=0.88) 

            self.overall_financial_summary_chart_canvas_widget = FigureCanvasTkAgg(self.overall_financial_summary_chart_figure, master=parent_frame)
            self.overall_financial_summary_chart_canvas_widget.draw()
            canvas_tk_widget = self.overall_financial_summary_chart_canvas_widget.get_tk_widget()
            canvas_tk_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True) 
            parent_frame.update_idletasks() 
            
            self.overall_charts_canvas.after_idle(lambda: self.overall_charts_canvas.configure(scrollregion=self.overall_charts_canvas.bbox("all")))

        except Exception as e:
            logging.error(f"ReportingTab: Error creating financial summary pie chart: {e}", exc_info=True)
            ttk.Label(parent_frame, text=f"Error creating financial chart: {e}").pack(expand=True, fill=tk.BOTH)
            self.overall_financial_summary_chart_figure = None 

    def display_all_stats(self):
        """Main function to refresh and display all statistics and charts."""
        logging.info("ReportingTab: Refreshing all statistics.")
        open_jobs_df, today = self._prepare_open_jobs_data() # Balance_numeric is outstanding
        source_df_for_intake = self.app.status_df 

        if open_jobs_df is None: 
            self.overall_stats_text_area.config(state=tk.NORMAL)
            self.overall_stats_text_area.delete('1.0', tk.END)
            self._insert_text_with_tags(self.overall_stats_text_area, "No data available in the application.", ("warning_text",))
            self.overall_stats_text_area.config(state=tk.DISABLED)
            # Clear charts as well
            if hasattr(self, 'overall_status_chart_canvas_widget') and self.overall_status_chart_canvas_widget:
                self.overall_status_chart_canvas_widget.get_tk_widget().destroy()
            for widget in self.overall_status_chart_frame.winfo_children(): widget.destroy()
            ttk.Label(self.overall_status_chart_frame, text="No data for status chart.").pack(expand=True, fill=tk.BOTH)
            
            if hasattr(self, 'overall_financial_summary_chart_canvas_widget') and self.overall_financial_summary_chart_canvas_widget:
                self.overall_financial_summary_chart_canvas_widget.get_tk_widget().destroy()
            for widget in self.overall_financial_summary_chart_frame.winfo_children(): widget.destroy()
            ttk.Label(self.overall_financial_summary_chart_frame, text="No data for financial summary chart.").pack(expand=True, fill=tk.BOTH)
            
            if hasattr(self, 'weekly_intake_chart_canvas_widget') and self.weekly_intake_chart_canvas_widget:
                self.weekly_intake_chart_canvas_widget.get_tk_widget().destroy()
            for widget in self.weekly_intake_chart_frame.winfo_children(): widget.destroy()
            ttk.Label(self.weekly_intake_chart_frame, text="No data for weekly intake chart.").pack(expand=True, fill=tk.BOTH)

            for pc_name_safe in list(self.coordinator_tabs_widgets.keys()):
                try:
                    tab_frame_to_forget = self.coordinator_tabs_widgets[pc_name_safe].master.master 
                    for i, tab_id in enumerate(self.stats_notebook.tabs()):
                        if self.stats_notebook.nametowidget(tab_id) == tab_frame_to_forget:
                            self.stats_notebook.forget(i)
                            break
                except Exception as e: logging.warning(f"ReportingTab: Error removing old tab for {pc_name_safe}: {e}")
                del self.coordinator_tabs_widgets[pc_name_safe]
            return

        num_total_jobs_loaded = len(self.app.status_df) if self.app.status_df is not None else 0
        self._populate_overall_pipeline_tab(open_jobs_df, today, num_total_jobs_loaded)
        self._create_status_distribution_chart(self.overall_status_chart_frame, open_jobs_df)
        self._create_financial_summary_chart(self.overall_financial_summary_chart_frame, open_jobs_df) 

        weekly_intake_data = self._prepare_weekly_intake_data(source_df_for_intake)
        self._create_weekly_intake_chart(self.weekly_intake_chart_frame, weekly_intake_data)

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
                except Exception as e: logging.warning(f"ReportingTab: Error removing old tab for {pc_name_safe}: {e}")
                del self.coordinator_tabs_widgets[pc_name_safe]
        
        logging.info("ReportingTab: Statistics refresh complete.")


    def _populate_overall_pipeline_tab(self, open_jobs_df, today, num_total_jobs_loaded):
        """
        Populates the text area of the 'Overall Pipeline Health' tab.
        Assumes 'Balance_numeric' in open_jobs_df is OUTSTANDING balance.
        """
        txt = self.overall_stats_text_area
        txt.config(state=tk.NORMAL)
        txt.delete('1.0', tk.END)
        
        num_open_jobs = len(open_jobs_df) if open_jobs_df is not None else 0

        self._insert_text_with_tags(txt, f"Overall Snapshot ({today.strftime('%Y-%m-%d %H:%M:%S')})", ("header",))
        self._insert_separator_line(txt) 
        txt.insert(tk.END, "Total Jobs in Current Dataset: ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_total_jobs_loaded}", ("bold_metric",))
        txt.insert(tk.END, "Currently Open Jobs (Active): ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_open_jobs}", ("bold_metric",))
        self._insert_text_with_tags(txt, "") 

        self._insert_text_with_tags(txt, "Open Job Status Counts (See Chart for Details):", ("subheader",))
        if open_jobs_df is not None and not open_jobs_df.empty:
            open_status_counts = open_jobs_df['Status'].value_counts()
            if not open_status_counts.empty:
                 summary_line = ", ".join([f"{status}: {count}" for status, count in open_status_counts.nlargest(3).items()])
                 if len(open_status_counts) > 3:
                     summary_line += ", ..."
                 self._insert_text_with_tags(txt, summary_line, ("indented_item",))
            else: self._insert_text_with_tags(txt, "No open jobs with status information.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "No open jobs found.", ("indented_item",))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Financial Summary (All Open Jobs - See Chart for Details):", ("subheader",))
        if open_jobs_df is not None and not open_jobs_df.empty and 'Balance_numeric' in open_jobs_df.columns and 'InvoiceTotal_numeric' in open_jobs_df.columns:
            total_invoice = open_jobs_df['InvoiceTotal_numeric'].sum()
            total_outstanding_balance = open_jobs_df['Balance_numeric'].sum() # Sum of 'Balance' (outstanding)
            total_collected_calculated = total_invoice - total_outstanding_balance

            txt.insert(tk.END, "Total Invoice Amount: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_invoice)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Collected (Calculated): ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_collected_calculated)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Remaining Balance (from 'Balance' column): ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_outstanding_balance)}", ("indented_item", "bold_metric"))
        elif open_jobs_df is not None and open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for financial summary.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Numeric financial columns not pre-calculated or available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
        
        self._insert_text_with_tags(txt, "Work-in-Progress Timing (All Open Jobs, from Turn-in Date):", ("subheader",))
        self._insert_separator_line(txt) 
        if open_jobs_df is not None and not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns:
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
                else: self._insert_text_with_tags(txt, "  Could not determine age distribution (Age_Bucket column missing or empty).", ("indented_item", "warning_text"))
            else: self._insert_text_with_tags(txt, "No valid job ages for timing summary.", ("indented_item", "warning_text"))
        elif open_jobs_df is not None and open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for timing statistics.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Timing columns (JobAge_days) not pre-calculated or available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "\"Stuck\" Jobs in Early Stages (Open Jobs > 3 Weeks in early status):", ("subheader",))
        self._insert_separator_line(txt) 
        if open_jobs_df is not None and not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns and 'Status' in open_jobs_df.columns:
            early_statuses = ["New", "Waiting Measure", "Ready to order"] 
            stuck_threshold_days = 21 
            stuck_jobs_df = open_jobs_df[(open_jobs_df['Status'].isin(early_statuses)) & (open_jobs_df['JobAge_days'] > stuck_threshold_days)]
            if not stuck_jobs_df.empty:
                self._insert_text_with_tags(txt, f"Found {len(stuck_jobs_df)} potentially stuck job(s):", ("indented_item", "warning_text"))
                for _, job in stuck_jobs_df.iterrows():
                    account_name_stuck = job.get('Account', 'N/A')
                    po_number_stuck = job.get('Invoice #', 'N/A') 
                    status_stuck = job.get('Status', 'N/A')
                    age_stuck = job.get('JobAge_days', 0)
                    pc_stuck = job.get('Project Coordinator', 'N/A')
                    line_stuck = (f"  - Account: {account_name_stuck} - PO #: {po_number_stuck}, "
                                  f"Status: {status_stuck}, Age: {age_stuck:.0f}d, PC: {pc_stuck}")
                    self._insert_text_with_tags(txt, line_stuck, ("indented_item",))
            else: self._insert_text_with_tags(txt, "No open jobs identified as \"stuck\" in early stages.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine stuck jobs (missing required data like JobAge_days or Status).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "High-Value Aging Jobs (Open Jobs > 8 Weeks & Outstanding Balance > $10,000):", ("subheader",))
        self._insert_separator_line(txt) 
        if open_jobs_df is not None and not open_jobs_df.empty and \
           'JobAge_days' in open_jobs_df.columns and \
           'Balance_numeric' in open_jobs_df.columns and \
           'InvoiceTotal_numeric' in open_jobs_df.columns: 
            aging_threshold_days = 56 
            value_threshold = 10000 
            # 'Balance_numeric' is outstanding balance
            hv_aging_df = open_jobs_df[
                (open_jobs_df['JobAge_days'] > aging_threshold_days) & 
                (open_jobs_df['Balance_numeric'] > value_threshold)
            ]
            if not hv_aging_df.empty:
                self._insert_text_with_tags(txt, f"Found {len(hv_aging_df)} high-value aging job(s) (outstanding balance > ${value_threshold:,.0f}):", ("indented_item", "warning_text"))
                for _, job in hv_aging_df.iterrows():
                    account_name = job.get('Account', 'N/A')
                    po_number = job.get('Invoice #', 'N/A') 
                    outstanding_balance_val = job.get('Balance_numeric', 0) # Directly from 'Balance' column
                    job_total_val = job.get('InvoiceTotal_numeric', 0) 
                    job_age_days = job.get('JobAge_days', 0)
                    project_coordinator = job.get('Project Coordinator', 'N/A')

                    outstanding_balance_str = self.app.CURRENCY_FORMAT.format(outstanding_balance_val)
                    job_total_str = self.app.CURRENCY_FORMAT.format(job_total_val) 
                    line = (f"  - {account_name} - PO #: {po_number}, " 
                            f"Outstanding Balance: {outstanding_balance_str} (Job Total: {job_total_str}), "
                            f"Age: {job_age_days:.0f}d, PC: {project_coordinator}")
                    self._insert_text_with_tags(txt, line, ("indented_item",))
            else: self._insert_text_with_tags(txt, "No high-value aging jobs identified with outstanding balance > $10,000.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine high-value aging jobs (missing required data like JobAge_days, Balance_numeric, or InvoiceTotal_numeric).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Total Financial Value by Age Bucket (Open Jobs):", ("subheader",))
        self._insert_separator_line(txt) 
        if open_jobs_df is not None and not open_jobs_df.empty and 'Age_Bucket' in open_jobs_df.columns and \
           'InvoiceTotal_numeric' in open_jobs_df.columns and 'Balance_numeric' in open_jobs_df.columns and \
           open_jobs_df['Age_Bucket'].notna().any() : 
            financial_by_bucket = open_jobs_df.groupby('Age_Bucket', observed=False).agg( 
                Total_Invoice_Value=('InvoiceTotal_numeric', 'sum'),
                Total_Outstanding_Balance=('Balance_numeric', 'sum'), # Sum of 'Balance' (outstanding)
                Job_Count=('Invoice #', 'count') 
            )
            if not financial_by_bucket.empty:
                for bucket, data in financial_by_bucket.iterrows():
                    calculated_collected = data['Total_Invoice_Value'] - data['Total_Outstanding_Balance']
                    txt.insert(tk.END, f"- {bucket} ({data['Job_Count']} jobs):\n", ("indented_item", "key_value_label"))
                    txt.insert(tk.END, f"  - Total Invoice Value: ", ("indented_item",))
                    self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(data['Total_Invoice_Value'])}", ("indented_item", "bold_metric"))
                    txt.insert(tk.END, f"  - Total Collected: ", ("indented_item",))
                    self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(calculated_collected)}", ("indented_item", "bold_metric"))
                    txt.insert(tk.END, f"  - Total Remaining Balance: ", ("indented_item",))
                    self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(data['Total_Outstanding_Balance'])}", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "No data to aggregate financial value by age bucket.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine financial value by age bucket (missing required data like Age_Bucket or financial columns).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        # --- START: New section for Jobs with Low Collection Percentage ---
        self._insert_text_with_tags(txt, "Jobs with Low Collection (Calculated Collected < 35% of Total):", ("subheader",))
        self._insert_separator_line(txt)

        if open_jobs_df is not None and not open_jobs_df.empty and \
           'Balance_numeric' in open_jobs_df.columns and \
           'InvoiceTotal_numeric' in open_jobs_df.columns and \
           'Status' in open_jobs_df.columns:

            statuses_to_exclude_for_low_collection = ['New', 'Waiting Measure']
            df_for_low_collection_check = open_jobs_df[
                ~open_jobs_df['Status'].isin(statuses_to_exclude_for_low_collection)
            ].copy()

            if not df_for_low_collection_check.empty:
                # Calculate ACTUAL collected amount and its percentage
                df_for_low_collection_check['Collected_Amount_Calculated'] = \
                    df_for_low_collection_check['InvoiceTotal_numeric'] - df_for_low_collection_check['Balance_numeric']
                
                df_for_low_collection_check['Collected_Percentage_Actual'] = 0.0 # Initialize
                non_zero_invoice_total_mask = df_for_low_collection_check['InvoiceTotal_numeric'] != 0
                
                df_for_low_collection_check.loc[non_zero_invoice_total_mask, 'Collected_Percentage_Actual'] = \
                    (df_for_low_collection_check.loc[non_zero_invoice_total_mask, 'Collected_Amount_Calculated'] / \
                     df_for_low_collection_check.loc[non_zero_invoice_total_mask, 'InvoiceTotal_numeric']) * 100
                
                low_collection_jobs_df = df_for_low_collection_check[
                    df_for_low_collection_check['Collected_Percentage_Actual'] < 35
                ]

                if not low_collection_jobs_df.empty:
                    self._insert_text_with_tags(txt, f"Found {len(low_collection_jobs_df)} job(s) with calculated collection below 35% in relevant statuses:", ("indented_item", "warning_text"))
                    for _, job_row in low_collection_jobs_df.iterrows():
                        inv_num = job_row.get('Invoice #', 'N/A')
                        acc_name = job_row.get('Account', 'N/A')
                        job_status = job_row.get('Status', 'N/A')
                        inv_total_val = job_row.get('InvoiceTotal_numeric', 0)
                        outstanding_bal_val = job_row.get('Balance_numeric', 0) # From 'Balance' column
                        collected_calc_val = job_row.get('Collected_Amount_Calculated', 0)
                        coll_perc = job_row.get('Collected_Percentage_Actual', 0)
                        pc = job_row.get('Project Coordinator', 'N/A')

                        inv_total_str = self.app.CURRENCY_FORMAT.format(inv_total_val)
                        outstanding_str = self.app.CURRENCY_FORMAT.format(outstanding_bal_val)
                        collected_calc_str = self.app.CURRENCY_FORMAT.format(collected_calc_val)
                        
                        job_info_line1 = f"{acc_name}: {job_status}"
                        job_info_line2 = (f"    Total: {inv_total_str}, Balance: {outstanding_str}, "
                                          f"Collected: {collected_calc_str} ({coll_perc:.1f}% of total)")
                        #job_info_line3 = f"    PC: {pc}"
                        
                        self._insert_text_with_tags(txt, job_info_line1, ("indented_item",))
                        txt.insert(tk.END, job_info_line2 + "\n\n", ("indented_item",))
                        #txt.insert(tk.END, job_info_line3 + "\n\n", ("indented_item",))
                else:
                    self._insert_text_with_tags(txt, "No jobs found with calculated collection below 35% in the specified relevant statuses.", ("indented_item",))
            else:
                self._insert_text_with_tags(txt, "No open jobs found in the relevant statuses to check for low collection.", ("indented_item",))
        else:
            self._insert_text_with_tags(txt, "Cannot determine jobs with low collection (missing required data like Balance, Invoice Total, or Status).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
        # --- END: New section ---

        txt.config(state=tk.DISABLED) 


    def _populate_coordinator_tab(self, pc_name_display, pc_open_jobs_df, today):
        """
        Populates the text area of a specific Project Coordinator's tab.
        Assumes 'Balance_numeric' in pc_open_jobs_df is OUTSTANDING balance.
        """
        pc_name_safe = str(pc_name_display).replace(".", "_dot_") 
        txt = self._create_or_get_coordinator_tab(pc_name_safe) 
        
        self._insert_text_with_tags(txt, f"Statistics for {pc_name_display} ({today.strftime('%Y-%m-%d %H:%M:%S')})", ("header",))
        self._insert_separator_line(txt) 
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
            pc_total_outstanding_balance = pc_open_jobs_df['Balance_numeric'].sum() # Sum of 'Balance' (outstanding)
            pc_total_collected_calculated = pc_total_invoice - pc_total_outstanding_balance

            txt.insert(tk.END, "  Total Invoice Amount: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_invoice)}",("indented_item", "bold_metric"))
            txt.insert(tk.END, "  Total Collected (Calculated): ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_collected_calculated)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "  Total Remaining Balance (from 'Balance' col): ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(pc_total_outstanding_balance)}", ("indented_item", "bold_metric"))
        elif pc_open_jobs_df.empty: self._insert_text_with_tags(txt, "  No open jobs for financial summary.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "  Financial columns (Balance_numeric, InvoiceTotal_numeric) not available.", ("indented_item", "warning_text"))
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
                else: self._insert_text_with_tags(txt, "    Could not determine age distribution for this coordinator.", ("indented_item", "warning_text"))
            else: self._insert_text_with_tags(txt, "  No valid job ages to distribute into buckets for this coordinator.", ("indented_item", "warning_text"))
            self._insert_text_with_tags(txt, "")

            if pc_open_jobs_df['JobAge_days'].notna().any():
                pc_total_job_days = pc_open_jobs_df['JobAge_days'].sum()
                txt.insert(tk.END, "Total Days in Progress (sum of job ages): ", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"{pc_total_job_days:.0f} days", ("indented_item", "bold_metric"))
            else: self._insert_text_with_tags(txt, "Total Days in Progress: N/A (No valid job ages)", ("indented_item", "warning_text"))
            self._insert_text_with_tags(txt, "")
        elif pc_open_jobs_df.empty: self._insert_text_with_tags(txt, "  No open jobs for timing statistics.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "  Timing data ('JobAge_days') not available for this coordinator.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
            
        txt.config(state=tk.DISABLED) 

    def on_tab_selected(self):
        """Called when the Reporting tab is selected in the main notebook."""
        logging.info("Reporting tab selected.")
        if self.app.status_df is None or self.app.status_df.empty:
             if self.overall_stats_text_area: 
                self.overall_stats_text_area.config(state=tk.NORMAL)
                self.overall_stats_text_area.delete('1.0', tk.END)
                self._insert_text_with_tags(self.overall_stats_text_area, "No data loaded in the 'Data Management' tab. Please load data first, then click 'Refresh All Statistics' on this tab.", ("warning_text",))
                self.overall_stats_text_area.config(state=tk.DISABLED)

             if self.weekly_intake_chart_frame:
                if hasattr(self, 'weekly_intake_chart_canvas_widget') and self.weekly_intake_chart_canvas_widget:
                    self.weekly_intake_chart_canvas_widget.get_tk_widget().destroy()
                    self.weekly_intake_chart_canvas_widget = None
                    self.weekly_intake_chart_figure = None
                for widget in self.weekly_intake_chart_frame.winfo_children(): widget.destroy()
                ttk.Label(self.weekly_intake_chart_frame, text="No data loaded. Refresh after loading data.").pack(expand=True, fill=tk.BOTH)


    # --- Methods for HTML Export (Phase 3) ---
    def get_formatted_text_content(self, section_key="overall"):
        """
        Extracts formatted text content from the specified text area.
        Args:
            section_key (str): "overall" for the main summary, or a project
                               coordinator's safe name for their specific tab.
        Returns:
            list: A list of (text_segment, list_of_applied_tkinter_tags) tuples.
                  Returns an empty list if the section is not found or has no content.
        """
        logging.debug(f"ReportingTab: get_formatted_text_content called for section_key: '{section_key}'")
        text_widget = None
        if section_key == "overall":
            text_widget = self.overall_stats_text_area
        else:
            text_widget = self.coordinator_tabs_widgets.get(section_key)

        if not text_widget or not text_widget.winfo_exists():
            logging.warning(f"ReportingTab: Text widget for section '{section_key}' not found or does not exist.")
            return []

        if text_widget.index(tk.END) == "1.0": 
             logging.info(f"ReportingTab: Text widget for section '{section_key}' is empty.")
             return []
        
        content_with_tags = []
        current_tags = set()
        
        try:
            dump_output = text_widget.dump("1.0", tk.END, text=True, tag=True)
            for key, value, index in dump_output:
                if key == "text":
                    if value: 
                        content_with_tags.append((value, sorted(list(current_tags))))
                elif key == "tagon":
                    current_tags.add(value)
                elif key == "tagoff":
                    current_tags.discard(value)
            logging.info(f"ReportingTab: Successfully extracted {len(content_with_tags)} text segments for section '{section_key}'.")
        except Exception as e:
            logging.error(f"ReportingTab: Error during text_widget.dump or processing for section '{section_key}': {e}", exc_info=True)
            return []
            
        return content_with_tags

    def save_chart_as_image(self, chart_key, output_image_path):
        """
        Saves the specified chart as an image file.
        Args:
            chart_key (str): Identifier for the chart (e.g., "overall_status_chart", 
                             "overall_financial_summary_chart", "weekly_intake_chart").
            output_image_path (str): The full path where the image should be saved.
        Returns:
            bool: True if the chart was saved successfully, False otherwise.
        """
        logging.debug(f"ReportingTab: save_chart_as_image called for chart_key: '{chart_key}' at path: '{output_image_path}'")
        figure_to_save = None

        if chart_key == "overall_status_chart":
            figure_to_save = self.overall_status_chart_figure
        elif chart_key == "overall_financial_summary_chart":
            figure_to_save = self.overall_financial_summary_chart_figure
        elif chart_key == "weekly_intake_chart": 
            figure_to_save = self.weekly_intake_chart_figure
        else:
            logging.warning(f"ReportingTab: Unknown chart_key '{chart_key}' for saving.")
            return False

        if figure_to_save is None:
            logging.warning(f"ReportingTab: Figure for chart_key '{chart_key}' is not available (None). Chart might not have been generated.")
            return False
        
        if not figure_to_save.get_axes(): 
            logging.warning(f"ReportingTab: Figure for chart_key '{chart_key}' has no axes. Cannot save an empty chart.")
            return False

        try:
            output_dir = os.path.dirname(output_image_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                logging.info(f"ReportingTab: Created directory for chart image: {output_dir}")

            figure_to_save.savefig(output_image_path, dpi=100, bbox_inches='tight')
            logging.info(f"ReportingTab: Chart '{chart_key}' saved successfully to '{output_image_path}'.")
            return True
        except Exception as e:
            logging.error(f"ReportingTab: Error saving chart '{chart_key}' to '{output_image_path}': {e}", exc_info=True)
            return False