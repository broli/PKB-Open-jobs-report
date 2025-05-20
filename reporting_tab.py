# reporting_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import logging

import config

# --- Matplotlib Imports ---
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt 
import matplotlib.colors as mcolors # For more color options

class ReportingTab(ttk.Frame):
    def __init__(self, parent_notebook, app_instance):
        super().__init__(parent_notebook)
        self.app = app_instance
        self.stats_notebook = None 
        self.overall_stats_text_area = None
        self.coordinator_tabs_widgets = {} 

        # --- For Matplotlib Charts ---
        self.overall_status_chart_canvas_widget = None 
        self.overall_status_chart_frame = None 
        self.overall_financial_summary_chart_canvas_widget = None 
        self.overall_financial_summary_chart_frame = None 
        
        self.overall_charts_canvas = None 
        self.overall_charts_scrollable_frame = None 

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
        text_widget.config(fg=default_fg)

    def _insert_text_with_tags(self, text_widget, text, tags=None):
        """Helper to insert text with specified tags into a Text widget."""
        processed_text = text
        if not isinstance(tags, tuple) and tags is not None: 
            tags = (tags,)
        text_widget.insert(tk.END, processed_text, tags)
        text_widget.insert(tk.END, "\n") 

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

        # --- Overall Pipeline Health Tab ---
        overall_tab_content_frame = ttk.Frame(self.stats_notebook, padding=(5,5))
        self.stats_notebook.add(overall_tab_content_frame, text="Overall Pipeline Health")
        
        # Container for the text area (left side)
        overall_text_container_frame = tk.Frame(overall_tab_content_frame) 
        overall_text_container_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5)) 

        report_bg = getattr(config, 'REPORT_SUB_TAB_BG_COLOR', 'white') 
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
        Returns:
            pd.DataFrame: The processed open jobs DataFrame.
            pd.Timestamp: The current timestamp (normalized to day).
        """
        if self.app.status_df is None or self.app.status_df.empty:
            logging.warning("ReportingTab: No status_df available for processing.")
            return None, pd.Timestamp.now().normalize()

        df_all_loaded = self.app.status_df.copy()
        today = pd.Timestamp.now().normalize() # For consistent age calculation

        # Define statuses to exclude for "open jobs"
        excluded_statuses = ['Closed', 'Cancelled/Postponed', config.REVIEW_MISSING_STATUS]
        open_jobs_df = df_all_loaded[~df_all_loaded['Status'].isin(excluded_statuses)].copy()

        if open_jobs_df.empty:
            logging.info("ReportingTab: No open jobs after filtering.")
            return open_jobs_df, today 

        # Convert currency columns to numeric, handling errors
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
                # Define age bins and labels for bucketing
                age_bins = [-1, 7, 21, 49, 56, float('inf')] # Bins define the right edge (e.g., -1 to 7 days)
                
                age_labels = [
                    'Job Age: 0-7 Days (First Week of Job Age)', 
                    'Job Age: 8-21 Days (Job is 2-3 Weeks Old)', 
                    'Job Age: 22-49 Days (Job is 4-7 Weeks Old)', 
                    'Job Age: 50-56 Days (Job is 8 Weeks Old)', 
                    'Job Age: Over 56 Days (Job is Older than 8 Weeks)'
                ]
                open_jobs_df['Age_Bucket'] = pd.cut(open_jobs_df['JobAge_days'], bins=age_bins, labels=age_labels, right=True)
            else:
                open_jobs_df['Age_Bucket'] = pd.NA # If no valid ages, bucket is also NA
        else:
            logging.warning("ReportingTab: 'Turn in Date' column not found. Cannot calculate job age.")
            open_jobs_df['JobAge_days'] = pd.NA
            open_jobs_df['Age_Bucket'] = pd.NA
            
        return open_jobs_df, today

    def _create_status_distribution_chart(self, parent_frame, open_jobs_df):
        """Creates and embeds a horizontal bar chart for status distribution."""
        # Clear previous chart if exists
        if hasattr(self, 'overall_status_chart_canvas_widget') and self.overall_status_chart_canvas_widget:
            self.overall_status_chart_canvas_widget.get_tk_widget().destroy()
            self.overall_status_chart_canvas_widget = None
        for widget in parent_frame.winfo_children(): widget.destroy() # Clear placeholder

        if open_jobs_df is None or open_jobs_df.empty or 'Status' not in open_jobs_df.columns:
            ttk.Label(parent_frame, text="No data for status chart.").pack(expand=True, fill=tk.BOTH)
            return
        
        status_counts = open_jobs_df['Status'].value_counts().sort_index()
        if status_counts.empty:
            ttk.Label(parent_frame, text="No status data to plot.").pack(expand=True, fill=tk.BOTH)
            return

        try:
            # Dynamically adjust figure height based on the number of statuses
            fig_height = max(3.5, len(status_counts) * 0.5) # Min height, scales with items
            fig = Figure(figsize=(6, fig_height), dpi=90) 
            ax = fig.add_subplot(111)
            
            # Use a pleasant color map
            colors = plt.cm.get_cmap('Pastel2', len(status_counts)) 
            bars = status_counts.plot(kind='barh', ax=ax, color=[colors(i) for i in range(len(status_counts))])
            
            ax.set_title('Open Jobs by Status', fontsize=config.DEFAULT_FONT_SIZE +1)
            ax.set_xlabel('Number of Open Jobs', fontsize=config.DEFAULT_FONT_SIZE)
            ax.set_ylabel('Status', fontsize=config.DEFAULT_FONT_SIZE)
            ax.tick_params(axis='both', which='major', labelsize=config.DEFAULT_FONT_SIZE - 1)
            
            ax.grid(axis='x', linestyle=':', alpha=0.7, color='gray') # Subtle grid
            ax.spines['top'].set_visible(False) 
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('darkgrey')
            ax.spines['bottom'].set_color('darkgrey')

            # Adjust layout to prevent labels from being cut off
            fig.subplots_adjust(left=0.30, right=0.95, top=0.90, bottom=0.15) 
            
            # Add value labels to bars
            for i, v in enumerate(status_counts):
                ax.text(v + 0.2, i, str(v), color='black', va='center', fontweight='normal', fontsize=config.DEFAULT_FONT_SIZE -1)
            
            # Embed chart in Tkinter
            self.overall_status_chart_canvas_widget = FigureCanvasTkAgg(fig, master=parent_frame)
            self.overall_status_chart_canvas_widget.draw()
            canvas_tk_widget = self.overall_status_chart_canvas_widget.get_tk_widget()
            canvas_tk_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True) 
            parent_frame.update_idletasks() 
            
            # Update scrollregion of the parent canvas after drawing
            self.overall_charts_canvas.after_idle(lambda: self.overall_charts_canvas.configure(scrollregion=self.overall_charts_canvas.bbox("all")))

        except Exception as e:
            logging.error(f"ReportingTab: Error creating status distribution chart: {e}", exc_info=True)
            ttk.Label(parent_frame, text=f"Error creating chart: {e}").pack(expand=True, fill=tk.BOTH)

    def _create_financial_summary_chart(self, parent_frame, open_jobs_df):
        """Creates and embeds a pie chart for the financial summary."""
        # Clear previous chart
        if hasattr(self, 'overall_financial_summary_chart_canvas_widget') and self.overall_financial_summary_chart_canvas_widget:
            self.overall_financial_summary_chart_canvas_widget.get_tk_widget().destroy()
            self.overall_financial_summary_chart_canvas_widget = None
        for widget in parent_frame.winfo_children(): widget.destroy()

        if open_jobs_df is None or open_jobs_df.empty or \
           'InvoiceTotal_numeric' not in open_jobs_df.columns or \
           'Balance_numeric' not in open_jobs_df.columns:
            ttk.Label(parent_frame, text="No data for financial summary chart.").pack(expand=True, fill=tk.BOTH)
            return

        total_invoice = open_jobs_df['InvoiceTotal_numeric'].sum()
        total_balance = open_jobs_df['Balance_numeric'].sum()
        total_collected = total_invoice - total_balance
        
        # Prepare data for pie chart
        pie_labels_for_legend = []
        pie_values = []
        pie_colors = []
        explode_values = [] 

        if total_collected > 0:
            pie_labels_for_legend.append(f'Collected (${total_collected:,.0f})')
            pie_values.append(total_collected)
            pie_colors.append('#77DD77') # Pastel green
            explode_values.append(0.03) 
        if total_balance > 0:
            pie_labels_for_legend.append(f'Remaining (${total_balance:,.0f})')
            pie_values.append(total_balance)
            pie_colors.append('#FFB347') # Pastel orange
            explode_values.append(0.03)

        if not pie_values or sum(pie_values) == 0:
             ttk.Label(parent_frame, text="Financial values are zero or not suitable for pie chart.").pack(expand=True, fill=tk.BOTH)
             return

        try:
            fig = Figure(figsize=(6, 4.5), dpi=90) 
            ax = fig.add_subplot(111) 
            
            wedges, texts, autotexts = ax.pie(
                pie_values, 
                autopct='%1.1f%%', 
                startangle=90, 
                colors=pie_colors,
                explode=explode_values,
                wedgeprops=dict(width=0.45, edgecolor='w'), # Donut-like effect
                pctdistance=0.75, 
                textprops={'fontsize': config.DEFAULT_FONT_SIZE - 1, 'color':'black', 'weight':'bold'}
            )
            
            ax.set_title(f'Open Jobs Financials\n(Total Invoice: ${total_invoice:,.0f})', fontsize=config.DEFAULT_FONT_SIZE, pad=15) 
            ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
            
            # Legend placement
            lgd = ax.legend(wedges, pie_labels_for_legend,
                      title="Breakdown",
                      loc="center left", 
                      bbox_to_anchor=(0.92, 0.5), 
                      fontsize=config.DEFAULT_FONT_SIZE - 1,
                      title_fontsize=config.DEFAULT_FONT_SIZE)

            # Adjust subplot to make room for the legend
            fig.subplots_adjust(left=0.05, bottom=0.05, right=0.70, top=0.88) 

            self.overall_financial_summary_chart_canvas_widget = FigureCanvasTkAgg(fig, master=parent_frame)
            self.overall_financial_summary_chart_canvas_widget.draw()
            canvas_tk_widget = self.overall_financial_summary_chart_canvas_widget.get_tk_widget()
            canvas_tk_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True) 
            parent_frame.update_idletasks() 
            
            self.overall_charts_canvas.after_idle(lambda: self.overall_charts_canvas.configure(scrollregion=self.overall_charts_canvas.bbox("all")))

        except Exception as e:
            logging.error(f"ReportingTab: Error creating financial summary pie chart: {e}", exc_info=True)
            ttk.Label(parent_frame, text=f"Error creating financial chart: {e}").pack(expand=True, fill=tk.BOTH)


    def display_all_stats(self):
        """Main function to refresh and display all statistics and charts."""
        logging.info("ReportingTab: Refreshing all statistics.")
        open_jobs_df, today = self._prepare_open_jobs_data()

        # Handle case where no data is loaded at all
        if open_jobs_df is None: 
            self.overall_stats_text_area.config(state=tk.NORMAL)
            self.overall_stats_text_area.delete('1.0', tk.END)
            self._insert_text_with_tags(self.overall_stats_text_area, "No data available in the application.", ("warning_text",))
            self.overall_stats_text_area.config(state=tk.DISABLED)
            
            # Clear charts
            if hasattr(self, 'overall_status_chart_canvas_widget') and self.overall_status_chart_canvas_widget:
                self.overall_status_chart_canvas_widget.get_tk_widget().destroy()
                self.overall_status_chart_canvas_widget = None
                for widget in self.overall_status_chart_frame.winfo_children(): widget.destroy()
                ttk.Label(self.overall_status_chart_frame, text="No data for status chart.").pack(expand=True, fill=tk.BOTH)
            
            if hasattr(self, 'overall_financial_summary_chart_canvas_widget') and self.overall_financial_summary_chart_canvas_widget:
                self.overall_financial_summary_chart_canvas_widget.get_tk_widget().destroy()
                self.overall_financial_summary_chart_canvas_widget = None
                for widget in self.overall_financial_summary_chart_frame.winfo_children(): widget.destroy()
                ttk.Label(self.overall_financial_summary_chart_frame, text="No data for financial summary chart.").pack(expand=True, fill=tk.BOTH)
            
            # Remove all coordinator tabs
            for pc_name_safe in list(self.coordinator_tabs_widgets.keys()):
                try:
                    # Find the tab in the notebook and forget it
                    tab_frame_to_forget = self.coordinator_tabs_widgets[pc_name_safe].master.master # pc_tab_frame
                    for i, tab_id in enumerate(self.stats_notebook.tabs()):
                        if self.stats_notebook.nametowidget(tab_id) == tab_frame_to_forget:
                            self.stats_notebook.forget(i)
                            break
                except Exception as e: logging.warning(f"ReportingTab: Error removing old tab for {pc_name_safe}: {e}")
                del self.coordinator_tabs_widgets[pc_name_safe]
            return

        # Populate Overall Pipeline Tab (text and charts)
        num_total_jobs_loaded = len(self.app.status_df) if self.app.status_df is not None else 0
        self._populate_overall_pipeline_tab(open_jobs_df, today, num_total_jobs_loaded)
        self._create_status_distribution_chart(self.overall_status_chart_frame, open_jobs_df)
        self._create_financial_summary_chart(self.overall_financial_summary_chart_frame, open_jobs_df) 

        # Populate/Update Project Coordinator Tabs
        project_coordinator_col = 'Project Coordinator'
        active_coordinators_safe = set() # Keep track of coordinators with current open jobs

        if project_coordinator_col in open_jobs_df.columns and not open_jobs_df.empty:
            # Get unique, non-null coordinators and sort them
            unique_coordinators = sorted([pc for pc in open_jobs_df[project_coordinator_col].unique() if pd.notna(pc)])
            
            for pc_name in unique_coordinators:
                pc_name_safe = str(pc_name).replace(".", "_dot_")
                active_coordinators_safe.add(pc_name_safe)
                pc_specific_df = open_jobs_df[open_jobs_df[project_coordinator_col] == pc_name].copy()
                self._populate_coordinator_tab(pc_name, pc_specific_df, today) 
        
        # Remove tabs for coordinators who no longer have open jobs
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
        """Populates the text area of the 'Overall Pipeline Health' tab."""
        txt = self.overall_stats_text_area
        txt.config(state=tk.NORMAL)
        txt.delete('1.0', tk.END)
        
        num_open_jobs = len(open_jobs_df) if open_jobs_df is not None else 0

        self._insert_text_with_tags(txt, f"Overall Snapshot ({today.strftime('%Y-%m-%d %H:%M:%S')})", ("header",))
        self._insert_text_with_tags(txt, "--------------------------------------------------")
        txt.insert(tk.END, "Total Jobs in Current Dataset: ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_total_jobs_loaded}", ("bold_metric",))
        txt.insert(tk.END, "Currently Open Jobs (Active): ", ("key_value_label",))
        self._insert_text_with_tags(txt, f"{num_open_jobs}", ("bold_metric",))
        self._insert_text_with_tags(txt, "") 

        self._insert_text_with_tags(txt, "Open Job Status Counts (See Chart for Details):", ("subheader",))
        if open_jobs_df is not None and not open_jobs_df.empty:
            open_status_counts = open_jobs_df['Status'].value_counts()
            if not open_status_counts.empty:
                 # Display top 3 statuses, then '...' if more
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
            total_balance = open_jobs_df['Balance_numeric'].sum()
            total_collected = total_invoice - total_balance
            txt.insert(tk.END, "Total Invoice Amount: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_invoice)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Collected: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_collected)}", ("indented_item", "bold_metric"))
            txt.insert(tk.END, "Total Remaining Balance: ", ("indented_item", "key_value_label"))
            self._insert_text_with_tags(txt, f"{self.app.CURRENCY_FORMAT.format(total_balance)}", ("indented_item", "bold_metric"))
        elif open_jobs_df is not None and open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for financial summary.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Numeric financial columns not pre-calculated or available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
        
        self._insert_text_with_tags(txt, "Work-in-Progress Timing (All Open Jobs, from Turn-in Date):", ("subheader",))
        self._insert_text_with_tags(txt, "-----------------------------------------------------------")
        if open_jobs_df is not None and not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns:
            if open_jobs_df['JobAge_days'].notna().any():
                avg_age = open_jobs_df['JobAge_days'].mean()
                txt.insert(tk.END, "Average Age (since turn-in): ", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"{avg_age:.2f} days", ("indented_item", "bold_metric"))
                
                # Oldest project details
                oldest_idx = open_jobs_df['JobAge_days'].idxmax()
                oldest = open_jobs_df.loc[oldest_idx]
                self._insert_text_with_tags(txt, "Oldest Project (since turn-in):", ("indented_item", "key_value_label"))
                self._insert_text_with_tags(txt, f"  - Invoice #: {oldest.get('Invoice #', 'N/A')}", ("indented_item", "bold_metric")) # Retained "Invoice #" here as it's a general section
                self._insert_text_with_tags(txt, f"  - Turn-in Date: {oldest['TurnInDate_dt'].strftime(config.DATE_FORMAT) if pd.notna(oldest['TurnInDate_dt']) else 'N/A'}", ("indented_item",))
                self._insert_text_with_tags(txt, f"  - Age: {oldest['JobAge_days']:.0f} days", ("indented_item", "bold_metric"))
                
                self._insert_text_with_tags(txt, "Open Job Age Distribution (since turn-in):", ("indented_item", "key_value_label"))
                if 'Age_Bucket' in open_jobs_df.columns and open_jobs_df['Age_Bucket'].notna().any():
                    bucket_counts = open_jobs_df['Age_Bucket'].value_counts().sort_index()
                    for bucket, count in bucket_counts.items():
                        txt.insert(tk.END, f"  - {bucket}: ", ("indented_item",)) # Bucket names are now the new labels
                        self._insert_text_with_tags(txt, f"{count} jobs", ("indented_item", "bold_metric"))
                else: self._insert_text_with_tags(txt, "  Could not determine age distribution (Age_Bucket column missing or empty).", ("indented_item", "warning_text"))
            else: self._insert_text_with_tags(txt, "No valid job ages for timing summary.", ("indented_item", "warning_text"))
        elif open_jobs_df is not None and open_jobs_df.empty: self._insert_text_with_tags(txt, "No open jobs for timing statistics.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Timing columns (JobAge_days) not pre-calculated or available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        # "Stuck" Jobs Section
        self._insert_text_with_tags(txt, "\"Stuck\" Jobs in Early Stages (Open Jobs > 3 Weeks in early status):", ("subheader",))
        self._insert_text_with_tags(txt, "---------------------------------------------------------------------")
        if open_jobs_df is not None and not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns and 'Status' in open_jobs_df.columns:
            early_statuses = ["New", "Waiting Measure", "Ready to order"] # Define "early" statuses
            stuck_threshold_days = 21 # More than 3 weeks
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

        # High-Value Aging Jobs Section
        # --- MODIFIED: Section title and value_threshold ---
        self._insert_text_with_tags(txt, "High-Value Aging Jobs (Open Jobs > 8 Weeks & Balance > $10,000):", ("subheader",))
        self._insert_text_with_tags(txt, "-------------------------------------------------------------------") # Adjusted underline
        if open_jobs_df is not None and not open_jobs_df.empty and 'JobAge_days' in open_jobs_df.columns and 'Balance_numeric' in open_jobs_df.columns and 'InvoiceTotal_numeric' in open_jobs_df.columns: 
            aging_threshold_days = 56 
            value_threshold = 10000 # --- MODIFIED: Increased threshold ---
            hv_aging_df = open_jobs_df[(open_jobs_df['JobAge_days'] > aging_threshold_days) & (open_jobs_df['Balance_numeric'] > value_threshold)]
            if not hv_aging_df.empty:
                self._insert_text_with_tags(txt, f"Found {len(hv_aging_df)} high-value aging job(s):", ("indented_item", "warning_text"))
                for _, job in hv_aging_df.iterrows():
                    account_name = job.get('Account', 'N/A')
                    po_number = job.get('Invoice #', 'N/A') 
                    remaining_balance_val = job.get('Balance_numeric', 0)
                    job_total_val = job.get('InvoiceTotal_numeric', 0) 
                    job_age_days = job.get('JobAge_days', 0)
                    project_coordinator = job.get('Project Coordinator', 'N/A')

                    remaining_balance_str = self.app.CURRENCY_FORMAT.format(remaining_balance_val)
                    job_total_str = self.app.CURRENCY_FORMAT.format(job_total_val) 

                    line = (f"  - {account_name} - PO #: {po_number}, " 
                            f"Remaining Balance: {remaining_balance_str} (Job Total: {job_total_str}), "
                            f"Age: {job_age_days:.0f}d, PC: {project_coordinator}")
                    self._insert_text_with_tags(txt, line, ("indented_item",))
            else: self._insert_text_with_tags(txt, "No high-value aging jobs identified.", ("indented_item",))
        else: self._insert_text_with_tags(txt, "Cannot determine high-value aging jobs (missing required data like JobAge_days, Balance_numeric, or InvoiceTotal_numeric).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        # Financial Value by Age Bucket Section
        self._insert_text_with_tags(txt, "Total Financial Value by Age Bucket (Open Jobs):", ("subheader",))
        self._insert_text_with_tags(txt, "------------------------------------------------")
        if open_jobs_df is not None and not open_jobs_df.empty and 'Age_Bucket' in open_jobs_df.columns and \
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
        else: self._insert_text_with_tags(txt, "Cannot determine financial value by age bucket (missing required data like Age_Bucket or financial columns).", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")
        txt.config(state=tk.DISABLED) # Set back to read-only


    def _populate_coordinator_tab(self, pc_name_display, pc_open_jobs_df, today):
        """Populates the text area of a specific Project Coordinator's tab."""
        pc_name_safe = str(pc_name_display).replace(".", "_dot_") # Sanitize for widget key
        txt = self._create_or_get_coordinator_tab(pc_name_safe) # Gets or creates the text widget
        
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
        else: self._insert_text_with_tags(txt, "  Financial columns (Balance_numeric, InvoiceTotal_numeric) not available.", ("indented_item", "warning_text"))
        self._insert_text_with_tags(txt, "")

        self._insert_text_with_tags(txt, "Work-in-Progress Timing (for these open jobs, from Turn-in Date):", ("subheader",))
        if not pc_open_jobs_df.empty and 'JobAge_days' in pc_open_jobs_df.columns:
            if 'Age_Bucket' in pc_open_jobs_df.columns and pc_open_jobs_df['Age_Bucket'].notna().any():
                self._insert_text_with_tags(txt, "Job Age Distribution:", ("indented_item", "key_value_label"))
                pc_bucket_counts = pc_open_jobs_df['Age_Bucket'].value_counts().sort_index()
                if not pc_bucket_counts.empty:
                    for bucket, count in pc_bucket_counts.items():
                        txt.insert(tk.END, f"    - {bucket}: ", ("indented_item",)) # Bucket names are the new labels
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
            
        txt.config(state=tk.DISABLED) # Set back to read-only

    def on_tab_selected(self):
        """Called when the Reporting tab is selected in the main notebook."""
        logging.info("Reporting tab selected.")
        # Check if data is loaded; if not, show a message.
        if self.app.status_df is None or self.app.status_df.empty:
             if self.overall_stats_text_area: # Ensure widget exists
                self.overall_stats_text_area.config(state=tk.NORMAL)
                self.overall_stats_text_area.delete('1.0', tk.END)
                self._insert_text_with_tags(self.overall_stats_text_area, "No data loaded in the 'Data Management' tab. Please load data first, then click 'Refresh All Statistics' on this tab.", ("warning_text",))
                self.overall_stats_text_area.config(state=tk.DISABLED)
        # Optionally, you could trigger a refresh here if desired,
        # but the button provides explicit control.
        # self.display_all_stats() 
