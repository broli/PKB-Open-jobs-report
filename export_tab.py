# export_tab.py
import tkinter as tk
from tkinter import ttk, messagebox 
import logging
import os # For path joining

import config # For accessing font, color configurations

class ExportTab(ttk.Frame):
    """
    Manages the UI and interactions for the Export Report tab.
    This tab will allow users to select parts of the report and export them to HTML.
    """
    def __init__(self, parent_notebook, app_instance):
        """
        Initialize the ExportTab.
        Args:
            parent_notebook: The ttk.Notebook widget this tab will belong to.
            app_instance: The instance of the main application (OpenJobsApp),
                          used to access shared data and methods.
        """
        super().__init__(parent_notebook)
        self.app = app_instance  # Store a reference to the main app
        
        # --- Variables for UI Controls (Phase 2) ---
        self.include_overall_health_var = tk.BooleanVar(value=True)
        self.include_coordinator_details_var = tk.BooleanVar(value=True)
        self.include_charts_var = tk.BooleanVar(value=True)
        # Optional: Date range variables can be added here if needed in the future
        # self.start_date_var = tk.StringVar() # Define if using the date entry fields below
        # self.end_date_var = tk.StringVar()   # Define if using the date entry fields below
        
        self._setup_ui() # Build the user interface for this tab

    def _setup_ui(self):
        """Creates and configures the widgets for this tab."""
        # Main frame for this tab
        main_frame = ttk.Frame(self, padding=(10, 10))
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Title Label
        title_label = ttk.Label(main_frame, text="Export Report to HTML", 
                                font=self.app.DEFAULT_FONT_BOLD if hasattr(self.app, 'DEFAULT_FONT_BOLD') else ('Calibri', 14, 'bold'))
        title_label.pack(pady=(0, 15))

        # --- Content Selection Controls (Phase 2) ---
        controls_frame = ttk.LabelFrame(main_frame, text="Content to Include", padding=(10, 10))
        controls_frame.pack(pady=10, padx=10, fill=tk.X, anchor=tk.N)

        # Checkbutton for Overall Pipeline Health
        overall_health_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include Overall Pipeline Health Summary", 
            variable=self.include_overall_health_var
        )
        overall_health_cb.pack(anchor=tk.W, pady=2)

        # Checkbutton for All Project Coordinator Details
        coordinator_details_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include All Project Coordinator Details", 
            variable=self.include_coordinator_details_var
        )
        coordinator_details_cb.pack(anchor=tk.W, pady=2)
        
        # Placeholder for future advanced coordinator selection (as per roadmap)
        # advanced_coordinator_label = ttk.Label(controls_frame, text=" (Advanced coordinator selection will be here later)")
        # advanced_coordinator_label.pack(anchor=tk.W, padx=(20,0), pady=1)


        # Checkbutton for Charts
        charts_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include Charts (Overall Status & Financial Summary)", 
            variable=self.include_charts_var
        )
        charts_cb.pack(anchor=tk.W, pady=2)

        # Optional: Date Range Entry Fields (can be uncommented and developed later)
        # date_range_frame = ttk.LabelFrame(main_frame, text="Filter by Date Range (Optional)", padding=(10,10))
        # date_range_frame.pack(pady=10, padx=10, fill=tk.X, anchor=tk.N)
        # ttk.Label(date_range_frame, text="Start Date (e.g., YYYY-MM-DD):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        # start_date_entry = ttk.Entry(date_range_frame, textvariable=self.start_date_var) # Requires self.start_date_var
        # start_date_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        # ttk.Label(date_range_frame, text="End Date (e.g., YYYY-MM-DD):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        # end_date_entry = ttk.Entry(date_range_frame, textvariable=self.end_date_var) # Requires self.end_date_var
        # end_date_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        # date_range_frame.columnconfigure(1, weight=1)

        # --- Export Button ---
        export_button_frame = ttk.Frame(main_frame) 
        export_button_frame.pack(pady=(20,10), fill=tk.X)

        # Storing the button in an instance variable if we need to change its state later
        self.export_button = ttk.Button(export_button_frame, text="Export to HTML", command=self._initiate_export_process, style="Accent.TButton")
        self.export_button.pack() 
        
        # Comment from Phase 2, explaining removal of Phase 1 placeholder:
        # Removed the placeholder label from Phase 1 as actual controls are now added.
        # controls_placeholder_label = ttk.Label(main_frame, text="\n(Future controls for content selection will appear here)")
        # controls_placeholder_label.pack(pady=10)


    def _tkinter_tag_to_css_class(self, tkinter_tag):
        """Maps a Tkinter tag name to a CSS class name."""
        tag_map = {
            "header": "report-header",
            "subheader": "report-subheader",
            "bold_metric": "report-bold-metric",
            "key_value_label": "report-key-value-label",
            "indented_item": "report-indented-item",
            "warning_text": "report-warning-text"
        }
        return tag_map.get(tkinter_tag, "")

    def _convert_tkinter_text_to_html(self, text_segment, tkinter_tags_list):
        """
        Converts a single text segment with its Tkinter tags into an HTML paragraph.
        Args:
            text_segment (str): The text content.
            tkinter_tags_list (list): A list of Tkinter tags applied to this segment.
        Returns:
            str: An HTML string (e.g., <p class="class1 class2">text</p>), or empty string.
        """
        if not text_segment and not tkinter_tags_list:
            return ""

        html_text = text_segment.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        
        if html_text == "\n":
            return "<br>\n" 
        
        html_text = html_text.strip("\n") 
        html_text = html_text.replace("\n", "<br>\n") 

        if not html_text.strip(): 
            return "" 

        css_classes = [self._tkinter_tag_to_css_class(tag) for tag in tkinter_tags_list if self._tkinter_tag_to_css_class(tag)]
        class_attribute = f' class="{ " ".join(css_classes) }"' if css_classes else ""
        
        return f'<p{class_attribute}>{html_text}</p>\n'

    def _generate_html_content(self, image_dir_full_path, image_subdir_name_for_html_src):
        """
        Generates the full HTML content string based on selected options.
        Args:
            image_dir_full_path (str): The full absolute path to the directory where
                                       chart images will be/are saved.
            image_subdir_name_for_html_src (str): The relative path/name of the image 
                                                  subdirectory to be used in HTML <img> src attributes.
        Returns:
            str: The complete HTML document as a string.
                 Returns None if critical components (like reporting_tab) are missing.
        """
        html_parts = []
        
        reporting_tab = self.app.reporting_tab_instance
        if not reporting_tab:
            logging.error("ExportTab: ReportingTab instance not found. Cannot generate HTML.")
            messagebox.showerror("Export Error", "Reporting data is not available. Cannot generate HTML.", parent=self)
            return None

        # --- HTML Header and CSS ---
        html_parts.append("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n")
        html_parts.append("    <meta charset=\"UTF-8\">\n")
        html_parts.append("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n")
        html_parts.append("    <title>Job Report</title>\n")
        html_parts.append("    <style>\n")
        html_parts.append("        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }\n")
        html_parts.append("        .report-container { max-width: 900px; margin: auto; padding: 20px; border: 1px solid #ddd; box-shadow: 0 0 10px rgba(0,0,0,0.1); }\n")
        html_parts.append("        h1, h2, h3 { color: #333; }\n")
        html_parts.append("        p { margin-bottom: 0.5em; margin-top: 0.2em; }\n") 
        html_parts.append("        .chart-image { max-width: 100%; height: auto; border: 1px solid #eee; margin-top: 10px; margin-bottom: 20px; }\n")
        
        default_fg = getattr(config, 'REPORT_TEXT_FG_COLOR', 'black')
        default_font_family = getattr(config, 'DEFAULT_FONT_FAMILY', 'Arial')
        default_font_size = getattr(config, 'DEFAULT_FONT_SIZE', 12)

        html_parts.append(f"        .report-header {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size + 4}pt; font-weight: bold; text-decoration: underline; color: #003366; margin-top: 1.5em; margin-bottom: 0.8em; }}\n")
        html_parts.append(f"        .report-subheader {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size + 2}pt; font-weight: bold; color: {default_fg}; margin-top: 1.2em; margin-bottom: 0.6em; }}\n")
        html_parts.append(f"        .report-bold-metric {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; font-weight: bold; color: {default_fg}; }}\n")
        html_parts.append(f"        .report-key-value-label {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: #404040; }}\n")
        html_parts.append(f"        .report-indented-item {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: {default_fg}; margin-left: 25px; }}\n")
        html_parts.append(f"        .report-warning-text {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: #990000; font-style: italic; }}\n")
        html_parts.append(f"        .coordinator-section {{ margin-top: 20px; padding-top: 15px; border-top: 1px dashed #ccc; }}\n")

        html_parts.append("    </style>\n</head>\n<body>\n<div class=\"report-container\">\n")
        html_parts.append("<h1>Open Jobs Report</h1>\n")

        # --- Process Selected Text Sections ---
        if self.include_overall_health_var.get():
            html_parts.append("<h2>Overall Pipeline Health</h2>\n")
            text_data_overall = reporting_tab.get_formatted_text_content("overall")
            if text_data_overall:
                for segment, tags in text_data_overall:
                    html_parts.append(self._convert_tkinter_text_to_html(segment, tags))
            else:
                html_parts.append("<p><em>No overall health data available or selected for export.</em></p>\n")
        
        if self.include_coordinator_details_var.get():
            html_parts.append("<div class=\"coordinator-section\">\n")
            html_parts.append("<h2>Project Coordinator Details</h2>\n")
            if reporting_tab.coordinator_tabs_widgets:
                coordinator_found = False
                sorted_coordinator_keys = sorted(reporting_tab.coordinator_tabs_widgets.keys())

                for pc_safe_name in sorted_coordinator_keys:
                    pc_display_name = pc_safe_name.replace("_dot_", ".") 
                    text_data_pc = reporting_tab.get_formatted_text_content(pc_safe_name)
                    if text_data_pc:
                        coordinator_found = True
                        html_parts.append(f"<h3>Coordinator: {pc_display_name}</h3>\n")
                        for segment, tags in text_data_pc:
                            html_parts.append(self._convert_tkinter_text_to_html(segment, tags))
                if not coordinator_found:
                     html_parts.append("<p><em>No specific coordinator data available for export.</em></p>\n")
            else:
                html_parts.append("<p><em>No project coordinator data available in the report.</em></p>\n")
            html_parts.append("</div>\n")

        # --- Process Selected Charts ---
        if self.include_charts_var.get():
            html_parts.append("<h2>Charts</h2>\n")
            charts_exported_count = 0

            status_chart_filename = "overall_status_chart.png"
            status_chart_full_path = os.path.join(image_dir_full_path, status_chart_filename)
            status_chart_html_src = os.path.join(image_subdir_name_for_html_src, status_chart_filename).replace("\\", "/")
            
            logging.info(f"Attempting to save status chart to: {status_chart_full_path}")
            if reporting_tab.save_chart_as_image("overall_status_chart", status_chart_full_path):
                html_parts.append(f'<div>\n<h3>Open Jobs by Status</h3>\n<img src="{status_chart_html_src}" alt="Overall Status Chart" class="chart-image">\n</div>\n')
                charts_exported_count += 1
            else:
                html_parts.append("<p><em>Overall status chart could not be exported.</em></p>\n")

            financial_chart_filename = "overall_financial_summary_chart.png"
            financial_chart_full_path = os.path.join(image_dir_full_path, financial_chart_filename)
            financial_chart_html_src = os.path.join(image_subdir_name_for_html_src, financial_chart_filename).replace("\\", "/")

            logging.info(f"Attempting to save financial chart to: {financial_chart_full_path}")
            if reporting_tab.save_chart_as_image("overall_financial_summary_chart", financial_chart_full_path):
                html_parts.append(f'<div>\n<h3>Financial Summary</h3>\n<img src="{financial_chart_html_src}" alt="Financial Summary Chart" class="chart-image">\n</div>\n')
                charts_exported_count += 1
            else:
                html_parts.append("<p><em>Financial summary chart could not be exported.</em></p>\n")
            
            if charts_exported_count == 0 and self.include_charts_var.get():
                 html_parts.append("<p><em>No charts were available or selected for export.</em></p>\n")

        # --- HTML Footer ---
        html_parts.append("</div>\n</body>\n</html>")
        
        return "".join(html_parts)

    def _initiate_export_process(self):
        """
        This method will eventually handle the full export process including
        file dialogs and writing to file (Phase 5).
        For Phase 4, it calls _generate_html_content for testing.
        """
        logging.info("Initiating HTML export process...")
        
        dummy_html_base_path = "." 
        image_subdir_name = "exported_report_images"
        image_dir_full_path = os.path.join(dummy_html_base_path, image_subdir_name)

        if not os.path.exists(image_dir_full_path):
            try:
                os.makedirs(image_dir_full_path, exist_ok=True)
                logging.info(f"Created dummy image directory for testing: {image_dir_full_path}")
            except Exception as e:
                logging.error(f"Could not create dummy image directory {image_dir_full_path}: {e}")
        
        html_content = self._generate_html_content(image_dir_full_path, image_subdir_name)

        if html_content:
            logging.info("HTML content generated successfully (length: {}).".format(len(html_content)))
            print("\n--- Generated HTML (Snippet) ---")
            print(html_content[:1000] + "...\n--- End Snippet ---\n")
            
            try:
                test_output_filename = "test_export.html"
                with open(test_output_filename, "w", encoding="utf-8") as f:
                    f.write(html_content)
                logging.info(f"Test HTML content written to {os.path.abspath(test_output_filename)}")
                messagebox.showinfo("HTML Generated (Test)", 
                                    f"HTML content generated and a test file saved as:\n{os.path.abspath(test_output_filename)}\n\n"
                                    f"Images (if any) were attempted to be saved in:\n{os.path.abspath(image_dir_full_path)}", 
                                    parent=self)
            except Exception as e:
                logging.error(f"Error writing test HTML file: {e}")
                messagebox.showerror("Test Save Error", f"Could not write test HTML file: {e}", parent=self)
        else:
            logging.warning("HTML content generation failed.")

    def on_tab_selected(self):
        """
        Called when this tab is selected in the notebook.
        """
        logging.info("Export Report tab selected.")
        reporting_tab = self.app.reporting_tab_instance
        if reporting_tab and reporting_tab.app.status_df is not None and not reporting_tab.app.status_df.empty:
            if hasattr(self, 'export_button'): 
                 self.export_button.config(state=tk.NORMAL)
        else:
            if hasattr(self, 'export_button'):
                 self.export_button.config(state=tk.DISABLED)
                 # Check if the tab is currently selected to avoid showing messagebox if it's not the active tab
                 # This requires a bit more complex check usually involving comparing self to notebook.select()
                 # For now, this messagebox might appear if data is cleared while this tab is not active,
                 # then selected. This could be refined if it becomes an issue.
                 # A flag could be set if messagebox already shown for this state.
                 # messagebox.showinfo("No Data", 
                 #                     "No data loaded in the 'Reporting & Statistics' tab. "
                 #                     "Please load data and refresh statistics before exporting.", 
                 #                     parent=self.app) 
        pass
