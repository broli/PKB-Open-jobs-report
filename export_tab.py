# export_tab.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
import os 
import pandas as pd
import webbrowser

import config 

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
        self.app = app_instance
        
        self.include_overall_health_var = tk.BooleanVar(value=True)
        self.include_coordinator_details_var = tk.BooleanVar(value=True)
        self.include_charts_var = tk.BooleanVar(value=True)
        
        self._setup_ui()

    def _setup_ui(self):
        """Creates and configures the widgets for this tab."""
        main_frame = ttk.Frame(self, padding=(10, 10))
        main_frame.pack(expand=True, fill=tk.BOTH)

        title_label = ttk.Label(main_frame, text="Export Report to HTML", 
                                font=self.app.DEFAULT_FONT_BOLD if hasattr(self.app, 'DEFAULT_FONT_BOLD') else ('Calibri', 14, 'bold'))
        title_label.pack(pady=(0, 15))

        controls_frame = ttk.LabelFrame(main_frame, text="Content to Include", padding=(10, 10))
        controls_frame.pack(pady=10, padx=10, fill=tk.X, anchor=tk.N)

        overall_health_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include Overall Pipeline Health Summary", 
            variable=self.include_overall_health_var
        )
        overall_health_cb.pack(anchor=tk.W, pady=2)

        coordinator_details_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include All Project Coordinator Details", 
            variable=self.include_coordinator_details_var
        )
        coordinator_details_cb.pack(anchor=tk.W, pady=2)
        
        charts_cb = ttk.Checkbutton(
            controls_frame, 
            text="Include Charts (Overall Status & Financial Summary)", 
            variable=self.include_charts_var
        )
        charts_cb.pack(anchor=tk.W, pady=2)

        export_button_frame = ttk.Frame(main_frame) 
        export_button_frame.pack(pady=(20,10), fill=tk.X)

        self.export_button = ttk.Button(export_button_frame, text="Export to HTML", command=self._initiate_export_process, style="Accent.TButton")
        self.export_button.pack() 
        
        self.on_tab_selected() 

    def _tkinter_tag_to_css_class(self, tkinter_tag):
        """Maps a Tkinter tag name to a CSS class name."""
        tag_map = {
            "header": "report-header",
            "subheader": "report-subheader",
            "bold_metric": "report-bold-metric",
            "key_value_label": "report-key-value-label",
            "indented_item": "report-indented-item",
            "warning_text": "report-warning-text",
            # "separator_line_tk" is handled directly in _convert_tkinter_text_to_html
        }
        return tag_map.get(tkinter_tag, f"tk-tag-{tkinter_tag}") 

    def _convert_tkinter_text_to_html(self, text_segment, tkinter_tags_list):
        """
        Converts a single text segment with its Tkinter tags into an HTML paragraph or span.
        Handles newlines by converting them to <br> or wrapping segments in <p>.
        Special handling for 'separator_line_tk' tag to convert to <hr>.
        """
        if "separator_line_tk" in tkinter_tags_list:
            return "<hr class=\"content-separator\">\n"

        if not text_segment and not tkinter_tags_list:
            return ""

        html_text = text_segment.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        
        if html_text == "\n": 
            return "<br>\n" 
        
        stripped_text = html_text.strip("\n") 
        processed_text = stripped_text.replace("\n", "<br>\n")

        if not processed_text.strip() and not tkinter_tags_list: 
            return ""

        css_classes = [self._tkinter_tag_to_css_class(tag) for tag in tkinter_tags_list if tag != "separator_line_tk" and self._tkinter_tag_to_css_class(tag)]
        class_attribute = f' class="{ " ".join(css_classes) }"' if css_classes else ""
        
        return f'<p{class_attribute}>{processed_text if processed_text.strip() else "&nbsp;"}</p>\n'


    def _generate_html_content(self, image_dir_full_path, image_subdir_name_for_html_src):
        """
        Generates the full HTML content string based on selected options, with improved CSS.
        Dates for data points will reflect config.DATE_FORMAT used in reporting_tab.
        Timestamps for report generation are formatted explicitly here.
        """
        html_parts = []
        
        reporting_tab = self.app.reporting_tab_instance
        if not reporting_tab:
            logging.error("ExportTab: ReportingTab instance not found. Cannot generate HTML.")
            messagebox.showerror("Export Error", "Reporting data is not available. Cannot generate HTML.", parent=self)
            return None

        default_font_family = getattr(config, 'DEFAULT_FONT_FAMILY', 'Arial')
        default_font_size = getattr(config, 'DEFAULT_FONT_SIZE', 11) 
        default_fg_color = getattr(config, 'REPORT_TEXT_FG_COLOR', '#333')

        html_parts.append("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n")
        html_parts.append("    <meta charset=\"UTF-8\">\n")
        html_parts.append("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n")
        # Report title timestamp
        html_parts.append(f"    <title>Job Report - {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}</title>\n")
        html_parts.append("    <style>\n")
        html_parts.append(f"        body {{ font-family: '{default_font_family}', 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: {default_font_size}pt; margin: 0; padding: 0; line-height: 1.6; color: {default_fg_color}; background-color: #f4f7f6; }}\n")
        html_parts.append("        .report-container { max-width: 960px; margin: 20px auto; padding: 20px 30px; background-color: #ffffff; border: 1px solid #dde2e1; border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }\n")
        html_parts.append("        h1, h2, h3 { color: #2c3e50; margin-top: 1.5em; margin-bottom: 0.7em; line-height: 1.3; }\n")
        html_parts.append("        h1 { font-size: 22pt; text-align: center; color: #1a5276; border-bottom: 2px solid #5dade2; padding-bottom: 0.4em; margin-bottom: 1em; font-weight: 600;}\n")
        html_parts.append("        h2 { font-size: 18pt; color: #1f618d; border-bottom: 1px solid #aed6f1; padding-bottom: 0.3em; margin-top: 2em; font-weight: 500;}\n")
        html_parts.append("        h3 { font-size: 15pt; color: #2980b9; margin-top: 1.8em; font-weight: 500;}\n")
        html_parts.append("        p { margin-top: 0.5em; margin-bottom: 0.5em; }\n")
        html_parts.append(f"        .report-header {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size + 3}pt; font-weight: bold; text-decoration: underline; color: #003366; margin-top: 1.5em; margin-bottom: 0.8em; }}\n")
        html_parts.append(f"        .report-subheader {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size + 2}pt; font-weight: bold; color: {default_fg_color}; margin-top: 1.2em; margin-bottom: 0.6em; border-bottom: 1px dotted #bdc3c7; padding-bottom: 0.2em; }}\n")
        html_parts.append(f"        .report-bold-metric {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; font-weight: bold; color: #2c3e50; }}\n")
        html_parts.append(f"        .report-key-value-label {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: #566573; }}\n")
        html_parts.append(f"        .report-indented-item {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: {default_fg_color}; margin-left: 25px; }}\n")
        html_parts.append(f"        .report-warning-text {{ font-family: '{default_font_family}', sans-serif; font-size: {default_font_size}pt; color: #c0392b; font-style: italic; font-weight: 500; }}\n")
        html_parts.append("        .charts-section { margin-top: 30px; padding-top: 20px; border-top: 2px solid #5dade2; }\n")
        html_parts.append("        .chart-container { margin-bottom: 30px; padding: 15px; background-color: #f9fafb; border: 1px solid #e5e8e8; border-radius: 4px; text-align: center; }\n")
        html_parts.append("        .chart-container h3 { margin-top: 0.5em; margin-bottom: 1em; font-size: 13pt; color: #34495e; }\n")
        html_parts.append("        .chart-image {{ max-width: 100%; height: auto; border: 1px solid #d5dbdb; border-radius: 4px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }}\n")
        html_parts.append("        .coordinator-section { margin-top: 30px; padding-top: 20px; border-top: 2px solid #5dade2; }\n")
        html_parts.append("        .coordinator-section h3 { border-bottom: none; margin-top: 1em; }\n") 
        html_parts.append("        hr.content-separator { border: 0; height: 1px; background-color: #ccc; margin: 1.5em 0; }\n")
        html_parts.append("        .footer-timestamp { text-align: center; margin-top: 40px; font-size: 0.85em; color: #7f8c8d; border-top: 1px solid #eaecee; padding-top: 15px; }\n")
        html_parts.append("    </style>\n</head>\n<body>\n<div class=\"report-container\">\n")
        html_parts.append(f"<h1>Open Jobs Report</h1>\n")

        if self.include_overall_health_var.get():
            html_parts.append("<h2>Overall Pipeline Health</h2>\n")
            text_data_overall = reporting_tab.get_formatted_text_content("overall")
            if text_data_overall:
                for segment, tags in text_data_overall:
                    html_parts.append(self._convert_tkinter_text_to_html(segment, tags))
            else:
                html_parts.append("<p class=\"report-indented-item\"><em>No overall health data available or selected for export.</em></p>\n")
        
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
                     html_parts.append("<p class=\"report-indented-item\"><em>No specific coordinator data available for export.</em></p>\n")
            else:
                html_parts.append("<p class=\"report-indented-item\"><em>No project coordinator data available in the report.</em></p>\n")
            html_parts.append("</div>\n")

        if self.include_charts_var.get():
            html_parts.append("<div class=\"charts-section\">\n")
            html_parts.append("<h2>Charts</h2>\n")
            charts_exported_count = 0

            status_chart_filename = "overall_status_chart.png"
            status_chart_full_path = os.path.join(image_dir_full_path, status_chart_filename)
            status_chart_html_src = os.path.join(image_subdir_name_for_html_src, status_chart_filename).replace("\\", "/")
            
            html_parts.append(f'<div class="chart-container">\n<h3>Open Jobs by Status</h3>\n')
            if reporting_tab.save_chart_as_image("overall_status_chart", status_chart_full_path):
                html_parts.append(f'<img src="{status_chart_html_src}" alt="Overall Status Chart" class="chart-image">\n')
                charts_exported_count += 1
            else:
                html_parts.append("<p class=\"report-warning-text\"><em>Overall status chart could not be exported.</em></p>\n")
            html_parts.append("</div>\n")

            financial_chart_filename = "overall_financial_summary_chart.png"
            financial_chart_full_path = os.path.join(image_dir_full_path, financial_chart_filename)
            financial_chart_html_src = os.path.join(image_subdir_name_for_html_src, financial_chart_filename).replace("\\", "/")

            html_parts.append(f'<div class="chart-container">\n<h3>Financial Summary</h3>\n')
            if reporting_tab.save_chart_as_image("overall_financial_summary_chart", financial_chart_full_path):
                html_parts.append(f'<img src="{financial_chart_html_src}" alt="Financial Summary Chart" class="chart-image">\n')
                charts_exported_count += 1
            else:
                html_parts.append("<p class=\"report-warning-text\"><em>Financial summary chart could not be exported.</em></p>\n")
            html_parts.append("</div>\n") 
            
            if charts_exported_count == 0 and self.include_charts_var.get(): 
                 html_parts.append("<p class=\"report-indented-item report-warning-text\"><em>No charts were generated or available for export.</em></p>\n")
            html_parts.append("</div>\n")

        # Report generated footer timestamp
        html_parts.append(f"<div class=\"footer-timestamp\"><p>Report Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}</p></div>\n")
        html_parts.append("</div>\n</body>\n</html>")
        
        return "".join(html_parts)

    def _initiate_export_process(self):
        """
        Handles the full export process.
        """
        logging.info("Initiating HTML export process...")

        default_filename = f"Job_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.html"
        html_file_path = filedialog.asksaveasfilename(
            title="Save HTML Report As",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialfile=default_filename,
            parent=self.app
        )

        if not html_file_path:
            logging.info("HTML export cancelled by user.")
            return

        image_dir_full_path = ""
        image_subdir_name_for_html_src = "" 

        if self.include_charts_var.get():
            html_file_basename = os.path.basename(html_file_path)
            html_file_name_without_ext, _ = os.path.splitext(html_file_basename)
            safe_subdir_name = "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in html_file_name_without_ext)
            image_subdir_name_for_html_src = f"{safe_subdir_name}_images"
            
            html_file_dir = os.path.dirname(html_file_path)
            image_dir_full_path = os.path.join(html_file_dir, image_subdir_name_for_html_src)
            
            try:
                if not os.path.exists(image_dir_full_path):
                    os.makedirs(image_dir_full_path, exist_ok=True)
                logging.info(f"Image directory prepared: {image_dir_full_path}")
            except OSError as e:
                logging.error(f"Error creating image directory '{image_dir_full_path}': {e}")
                messagebox.showerror("Export Error", f"Could not create image directory:\n{image_dir_full_path}\n\nError: {e}", parent=self)
                return

        html_content = self._generate_html_content(image_dir_full_path, image_subdir_name_for_html_src)

        if not html_content:
            logging.warning("HTML content generation failed. Aborting export.")
            return

        try:
            with open(html_file_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            logging.info(f"HTML report successfully saved to: {html_file_path}")
            
            success_message = f"Report exported successfully to:\n{html_file_path}"
            if self.include_charts_var.get() and image_dir_full_path and os.path.exists(image_dir_full_path):
                success_message += f"\n\nChart images (if any) saved in subdirectory:\n{image_dir_full_path}"
            messagebox.showinfo("Export Successful", success_message, parent=self)
            
            try:
                url = 'file://' + os.path.abspath(html_file_path)
                webbrowser.open(url, new=2) 
                logging.info(f"Attempted to open '{url}' in web browser.")
            except Exception as e_open:
                logging.error(f"Error attempting to open HTML file in browser: {e_open}", exc_info=True)
                messagebox.showwarning("Open Warning", 
                                       f"Report exported, but could not automatically open it.\nError: {e_open}\n\nPlease open manually:\n{html_file_path}", 
                                       parent=self)

        except IOError as e:
            logging.error(f"Error writing HTML file to '{html_file_path}': {e}")
            messagebox.showerror("Export Error", f"Could not write HTML file:\n{html_file_path}\n\nError: {e}", parent=self)
        except Exception as e:
            logging.error(f"An unexpected error occurred during HTML file writing: {e}", exc_info=True)
            messagebox.showerror("Export Error", f"An unexpected error occurred while saving the HTML file: {e}", parent=self)

    def on_tab_selected(self):
        """
        Called when this tab is selected in the notebook.
        Updates the state of the export button based on data availability.
        """
        logging.info("Export Report tab selected.")
        try:
            reporting_tab = self.app.reporting_tab_instance
            if reporting_tab and hasattr(reporting_tab, 'app') and \
               reporting_tab.app.status_df is not None and not reporting_tab.app.status_df.empty:
                if hasattr(self, 'export_button'): 
                     self.export_button.config(state=tk.NORMAL)
            else:
                if hasattr(self, 'export_button'):
                     self.export_button.config(state=tk.DISABLED)
        except AttributeError:
            if hasattr(self, 'export_button'):
                self.export_button.config(state=tk.DISABLED)
            logging.warning("ExportTab: Could not determine reporting_tab status during on_tab_selected, disabling export button.")
        pass