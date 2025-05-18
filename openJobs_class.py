# https://github.com/broli/PKB-Open-jobs-report
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
    "Cancelled/Postponed", "New", "Closed",
    "Review - Missing from Report" # <-- New status added
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
    'Status': 200, # Increased width a bit for the new longer status
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
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

        # --- Global Font Configuration (Applied BEFORE WIDGET CREATION) ---
        default_font_string = f"{DEFAULT_FONT_FAMILY} {DEFAULT_FONT_SIZE}"
        self.option_add("*Font", default_font_string)
        self.option_add("*Text*Font", default_font_string)
        self.option_add("*Label*Font", default_font_string)
        self.option_add("*Button*Font", default_font_string)
        self.option_add("*Menu*Font", default_font_string)
        self.option_add("*MenuItem*Font", default_font_string)
        self.option_add("*TCombobox*Listbox*Font", default_font_string)

        self.maximize_window()

        # --- TTK Style Configuration ---
        self.style = ttk.Style(self)
        self.style.configure('.', font=DEFAULT_FONT)
        self.style.configure("Treeview", font=DEFAULT_FONT, rowheight=30)
        self.style.configure("Treeview.Heading", font=DEFAULT_FONT_BOLD)
        self.style.configure("TCombobox", font=DEFAULT_FONT)

        # Define colors for different row statuses in the Treeview
        self.status_colors = {
            "default_fg": "black", "default_bg": "white",
            "action_needed_bg": "#FFEBEE", "action_needed_fg": "black", # Light Red
            "all_good_bg": "#E8F5E9", "all_good_fg": "black",        # Light Green
            "closed_bg": "#E0E0E0", "closed_fg": "#757575",          # Light Grey
            "new_bg": "#E3F2FD", "new_fg": "black",                  # Light Blue
            "review_missing_bg": "#FFF9C4", "review_missing_fg": "#795548", # Light Yellow, Brown text <-- New color
            "selected_bg": "#B0BEC5", "selected_fg": "black"         # Blue Grey
        }

        self.style.map("Treeview",
                       background=[('selected', self.status_colors["selected_bg"])],
                       foreground=[('selected', self.status_colors["selected_fg"])])

        # --- Data Loading ---
        self.status_df = load_status()
        if self.status_df is None:
            messagebox.showerror("Fatal Error", "Could not load or initialize status data. Exiting.")
            self.destroy()
            return

        missing_cols = [col for col in EXPECTED_COLUMNS if col not in self.status_df.columns]
        if missing_cols:
            messagebox.showwarning("Data Warning",
                                   f"Loaded data is missing columns: {', '.join(missing_cols)}. "
                                   f"They will be added with default values (None).")
            for col in missing_cols:
                self.status_df[col] = None
        self.status_df = self.status_df[EXPECTED_COLUMNS]

        # --- UI Creation ---
        self.create_widgets()
        self.editing_window = None
        self.populate_treeview()

    def maximize_window(self):
        """Attempts to maximize the application window to fill the screen."""
        try:
            self.state('zoomed')
        except tk.TclError:
            try:
                m = self.maxsize()
                self.geometry('{}x{}+0+0'.format(*m))
            except tk.TclError:
                self.attributes('-fullscreen', True)

    def create_widgets(self):
        """Creates and configures the main UI widgets of the application."""
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel)
        file_menu.add_command(label="Save Current Status", command=self.save_data)
        file_menu.add_command(label="Generate Open Jobs Report", command=self.generate_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit_app)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menu_bar)

        self.tree = ttk.Treeview(self, columns=EXPECTED_COLUMNS, show="headings")
        for col in EXPECTED_COLUMNS:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected_row)


    def configure_treeview_columns(self):
        """
        Ensures the Treeview columns match the current DataFrame's columns.
        """
        current_tree_cols = list(self.status_df.columns)
        self.tree.configure(columns=current_tree_cols)
        for col in current_tree_cols:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)
        if '#0' in self.tree.column('#0'):
            self.tree.column('#0', width=0, stretch=tk.NO)

    def populate_treeview(self):
        """
        Clears and repopulates the Treeview with data from self.status_df.
        """
        for i in self.tree.get_children():
            self.tree.delete(i)

        if self.status_df is None or self.status_df.empty:
            logging.info("No data to populate in the treeview.")
            return

        self.status_df = self.status_df[EXPECTED_COLUMNS]
        date_columns = ['Order Date', 'Turn in Date']

        for df_index, row in self.status_df.iterrows():
            values = []
            for col_name in EXPECTED_COLUMNS:
                value = row[col_name]
                if col_name in date_columns and pd.notna(value):
                    try:
                        value = pd.to_datetime(value).strftime(DATE_FORMAT)
                    except (ValueError, TypeError):
                        logging.warning(f"Could not format date for '{value}' in column '{col_name}'. Original value used.")
                elif col_name in CURRENCY_COLUMNS and pd.notna(value):
                    try:
                        num_value = float(str(value).replace('$', '').replace(',', ''))
                        value = CURRENCY_FORMAT.format(num_value)
                    except (ValueError, TypeError):
                        logging.warning(f"Could not format currency for '{value}' in column '{col_name}'. Original value used.")
                values.append(value if pd.notna(value) else "")
            self.tree.insert("", tk.END, values=tuple(values), tags=(str(df_index),))

        self.after(10, self.set_column_widths_from_preferred)
        self.color_rows()

    def load_new_excel(self):
        """
        Handles loading new Excel data, processing it, and updating the Treeview.
        """
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if not excel_file_path:
            return

        new_df_raw = load_excel(excel_file_path)
        if new_df_raw is None:
            return

        current_status_df_copy = self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=EXPECTED_COLUMNS)
        self.status_df = process_data(new_df_raw, current_status_df_copy)

        if self.status_df is None:
            messagebox.showerror("Error", "Failed to process the new Excel data. Reverting to previous data.")
            self.status_df = current_status_df_copy
            return

        self.status_df = self.status_df[EXPECTED_COLUMNS]
        self.configure_treeview_columns()
        self.populate_treeview()
        messagebox.showinfo("Success", "New Excel data loaded and processed.")

    def save_data(self):
        """Saves the current status DataFrame."""
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.")
            return
        save_status(self.status_df)

    def generate_report(self):
        """Generates an Excel report of open invoices."""
        if self.status_df is None or self.status_df.empty:
            messagebox.showinfo("Info", "No data available to generate a report.")
            return

        open_invoices_df = self.status_df[
            ~self.status_df['Status'].isin(['Closed', 'Cancelled/Postponed', "Review - Missing from Report"]) # Exclude review status too
        ].copy()
        open_invoices_df = open_invoices_df[EXPECTED_COLUMNS]

        if open_invoices_df.empty:
            messagebox.showinfo("Info", "No open invoices to report.")
            return

        try:
            report_file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Open Invoices Report As",
                initialfile=OUTPUT_FILE
            )
            if not report_file_path:
                messagebox.showinfo("Info", "Report generation cancelled.")
                return

            with pd.ExcelWriter(report_file_path, engine='xlsxwriter') as writer:
                open_invoices_df.to_excel(writer, index=False, sheet_name='Open Invoices')
                workbook = writer.book
                worksheet = writer.sheets['Open Invoices']
                currency_format_excel = workbook.add_format({'num_format': '$#,##0.00'})
                for col_name in CURRENCY_COLUMNS:
                    if col_name in open_invoices_df.columns:
                        col_idx = open_invoices_df.columns.get_loc(col_name)
                        col_letter = chr(ord('A') + col_idx)
                        worksheet.set_column(f'{col_letter}:{col_letter}', None, currency_format_excel)
                for i, col in enumerate(open_invoices_df.columns):
                    column_len = max(open_invoices_df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len)
            messagebox.showinfo("Success", f"Report generated successfully: {report_file_path}")
        except Exception as e:
            logging.error(f"Error generating report: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred while generating the report: {e}")

    def on_double_click(self, event):
        """Handles double-click events for editing Status or Notes."""
        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)

        if not item_id or not column_id_str:
            return

        try:
            column_index_tree = int(column_id_str.replace("#", "")) - 1
            if not (0 <= column_index_tree < len(EXPECTED_COLUMNS)):
                logging.warning(f"Invalid column index from tree: {column_index_tree}")
                return
            actual_column_name = EXPECTED_COLUMNS[column_index_tree]
            df_row_index_str = self.tree.item(item_id, "tags")[0]
            df_row_index = int(df_row_index_str)

            if self.editing_window and self.editing_window.winfo_exists():
                self.editing_window.destroy()
            self.editing_window = None

            if actual_column_name == "Status":
                self.create_status_editor(item_id, df_row_index, actual_column_name)
            elif actual_column_name == "Notes":
                self.create_notes_editor(item_id, df_row_index, actual_column_name)
            else:
                logging.debug(f"No special editor for column '{actual_column_name}'. Double-click ignored.")
        except (ValueError, IndexError, TypeError) as e:
            logging.error(f"Error in on_double_click: {e}. Item ID: {item_id}, Column ID Str: {column_id_str}", exc_info=True)

    def _common_editor_save(self, item_id, df_row_index, column_name, new_value, editor_window):
        """Helper function to save edited data from popup editors."""
        try:
            self.status_df.loc[df_row_index, column_name] = new_value
            current_tree_values = list(self.tree.item(item_id, "values"))
            if column_name in EXPECTED_COLUMNS:
                column_tree_idx = EXPECTED_COLUMNS.index(column_name)
                display_value = new_value
                if column_name in CURRENCY_COLUMNS and pd.notna(new_value):
                    try: display_value = CURRENCY_FORMAT.format(float(str(new_value).replace('$', '').replace(',', '')))
                    except: pass
                elif column_name in ['Order Date', 'Turn in Date'] and pd.notna(new_value):
                    try: display_value = pd.to_datetime(new_value).strftime(DATE_FORMAT)
                    except: pass
                current_tree_values[column_tree_idx] = display_value
                self.tree.item(item_id, values=tuple(current_tree_values))
            else:
                logging.error(f"Column '{column_name}' not in EXPECTED_COLUMNS during save.")

            editor_window.destroy()
            self.editing_window = None
            self.color_rows() # Re-color as status might have changed
        except Exception as e:
            logging.error(f"Error in _common_editor_save for column '{column_name}': {e}", exc_info=True)
            messagebox.showerror("Save Error", f"Could not save change for {column_name}: {e}")

    def create_notes_editor(self, item_id, df_row_index, actual_column_name):
        """Creates a Toplevel window to edit the 'Notes' field."""
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {actual_column_name}")
        self.editing_window.transient(self); self.editing_window.grab_set()

        text_widget = tk.Text(self.editing_window, width=60, height=10, wrap=tk.WORD, font=DEFAULT_FONT)
        current_value = str(self.status_df.loc[df_row_index, actual_column_name])
        text_widget.insert(tk.END, current_value)
        text_widget.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING, fill=tk.BOTH, expand=True)
        text_widget.focus()

        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(0,DEFAULT_PADDING), padx=DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._common_editor_save(item_id, df_row_index, actual_column_name,
                                                                      text_widget.get("1.0", tk.END).strip(),
                                                                      self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.center_toplevel(self.editing_window)
        self.wait_window(self.editing_window)

    def create_status_editor(self, item_id, df_row_index, column_name):
        """Creates a Toplevel window to edit the 'Status' field."""
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self); self.editing_window.grab_set()

        current_value = str(self.status_df.loc[df_row_index, column_name])
        status_var = StringVar(self.editing_window)
        if current_value in ALLOWED_STATUS:
            status_var.set(current_value)
        elif ALLOWED_STATUS:
            status_var.set(ALLOWED_STATUS[0]) # Default if current invalid
        else:
            status_var.set("")

        inv_num_col = 'Invoice #'
        inv_num = self.status_df.loc[df_row_index, inv_num_col] if inv_num_col in self.status_df.columns else "N/A"
        ttk.Label(self.editing_window, text=f"Status for Invoice {inv_num}:").pack(padx=DEFAULT_PADDING,pady=(DEFAULT_PADDING,5))

        combo = ttk.Combobox(self.editing_window, textvariable=status_var, values=ALLOWED_STATUS, state="readonly", font=DEFAULT_FONT) # Ensure font for Combobox entry too
        combo.pack(padx=DEFAULT_PADDING, pady=5, fill=tk.X); combo.focus()

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
        """Deletes selected row(s) from DataFrame and refreshes Treeview."""
        selected_tree_items = self.tree.selection()
        if not selected_tree_items:
            messagebox.showinfo("No Selection", "Please select one or more rows to delete.")
            return

        confirm_msg = f"Are you sure you want to delete {len(selected_tree_items)} selected row(s)? This action cannot be undone from the UI."
        if not messagebox.askyesno("Confirm Delete", confirm_msg):
            return

        df_indices_to_delete = sorted([int(self.tree.item(item_id, "tags")[0]) for item_id in selected_tree_items], reverse=True)

        if df_indices_to_delete:
            self.status_df.drop(index=df_indices_to_delete, inplace=True)
            self.status_df.reset_index(drop=True, inplace=True)

        self.populate_treeview()
        messagebox.showinfo("Success", f"{len(selected_tree_items)} row(s) deleted.")

    def set_column_widths_from_preferred(self):
        """Sets Treeview column widths based on PREFERRED_COLUMN_WIDTHS."""
        self.update_idletasks()
        for col_name in EXPECTED_COLUMNS:
            if col_name == '#0': continue
            width = PREFERRED_COLUMN_WIDTHS.get(col_name, 100)
            final_width = max(MIN_COLUMN_WIDTH, int(width))
            final_width = min(MAX_COLUMN_WIDTH, int(final_width))
            self.tree.column(col_name, width=final_width, anchor=tk.W)
        if '#0' in self.tree['columns']: # Should not be in EXPECTED_COLUMNS
            self.tree.column('#0', width=0, stretch=tk.NO)

    def color_rows(self):
        """Applies background and foreground colors to Treeview rows based on 'Status'."""
        if self.status_df is None or self.status_df.empty:
            return

        styles_map = {
            "default_status_style": (self.status_colors["default_bg"], self.status_colors["default_fg"]),
            "action_needed_style": (self.status_colors["action_needed_bg"], self.status_colors["action_needed_fg"]),
            "all_good_style": (self.status_colors["all_good_bg"], self.status_colors["all_good_fg"]),
            "closed_style": (self.status_colors["closed_bg"], self.status_colors["closed_fg"]),
            "new_style": (self.status_colors["new_bg"], self.status_colors["new_fg"]),
            "review_missing_style": (self.status_colors["review_missing_bg"], self.status_colors["review_missing_fg"]) # <-- New style
        }
        for tag_name, (bg, fg) in styles_map.items():
            self.tree.tag_configure(tag_name, background=bg, foreground=fg)

        action_statuses = ["Ready to order", "Permit", "Waiting Measure"]
        good_statuses = ["Ready to dispatch", "In install", "Done", "Waiting for materials"]
        # 'New', 'Closed', 'Cancelled/Postponed', "Review - Missing from Report" have specific styles

        status_col_name = "Status"
        if status_col_name not in EXPECTED_COLUMNS:
            logging.error(f"'{status_col_name}' column missing. Cannot color rows accurately.")
            return

        try:
            status_column_tree_index = EXPECTED_COLUMNS.index(status_col_name)
        except ValueError:
            logging.error(f"'{status_col_name}' column not found in Treeview columns. Cannot color rows.")
            return

        for item_id in self.tree.get_children():
            try:
                df_index_tag = self.tree.item(item_id, "tags")[0]
                new_tags_for_item = [df_index_tag]

                values = self.tree.item(item_id, "values")
                if values and len(values) > status_column_tree_index:
                    status = str(values[status_column_tree_index])

                    if status == "New":
                        new_tags_for_item.append("new_style")
                    elif status == "Closed" or status == "Cancelled/Postponed":
                        new_tags_for_item.append("closed_style")
                    elif status == "Review - Missing from Report": # <-- Logic for new status
                        new_tags_for_item.append("review_missing_style")
                    elif status in action_statuses:
                        new_tags_for_item.append("action_needed_style")
                    elif status in good_statuses:
                        new_tags_for_item.append("all_good_style")
                    else:
                        new_tags_for_item.append("default_status_style")
                else:
                    new_tags_for_item.append("default_status_style")
                self.tree.item(item_id, tags=tuple(new_tags_for_item))
            except Exception as e:
                logging.error(f"Error coloring row {item_id}: {e}", exc_info=True)
                df_idx_tag_fallback = self.tree.item(item_id,"tags")[0] if self.tree.item(item_id,"tags") else "err_idx_fallback"
                self.tree.item(item_id, tags=(df_idx_tag_fallback, "default_status_style"))


    def sort_treeview_column(self, col, reverse):
        """Sorts the Treeview rows based on the clicked column."""
        if self.status_df is None or self.status_df.empty:
            return

        date_columns = ['Order Date', 'Turn in Date'] # Define which columns are dates

        try:
            items_to_sort = []
            for item_id in self.tree.get_children(''):
                sort_value = None
                if col in date_columns:
                    try:
                        # For date columns, get the original data from the DataFrame
                        df_index_tag = self.tree.item(item_id, "tags")[0]
                        df_index = int(df_index_tag)
                        original_value = self.status_df.loc[df_index, col]
                        # Convert to datetime, coercing errors to NaT (Not a Time)
                        # This handles actual datetime objects, parseable strings, None, or NaT already.
                        sort_value = pd.to_datetime(original_value, errors='coerce')
                    except (IndexError, KeyError, ValueError, TypeError) as e:
                        logging.warning(f"Could not get original date for sorting column {col}, item {item_id}: {e}. Using displayed value as fallback.")
                        # Fallback: use the displayed value from the tree if original data access fails
                        sort_value = self.tree.set(item_id, col)
                else:
                    # For non-date columns, use the displayed value from the tree
                    sort_value = self.tree.set(item_id, col)
                
                items_to_sort.append((sort_value, item_id))

            # Define a key function for sorting
            # This key returns tuples to allow for stable sorting across different data types
            def sort_key(item_tuple):
                value = item_tuple[0]  # This is the sort_value determined above

                if isinstance(value, pd.Timestamp): # Actual datetime objects
                    return (0, value)  # Sort first by type (0), then by timestamp value
                
                if pd.isna(value):  # Handles pd.NaT (from dates) or None (possibly from other fallbacks)
                    # Group NaT/None values. Place them consistently at one end.
                    # Using a fixed early/late timestamp helps group them.
                    # (1, ...) ensures they sort after valid Timestamps if Timestamp.min/max is used.
                    return (1, pd.Timestamp.min if not reverse else pd.Timestamp.max)

                # If not a Timestamp or NaT, it's likely a string (from other columns or date fallback)
                str_value = str(value) # Ensure it's a string for the following operations

                # Handle currency columns
                if col in CURRENCY_COLUMNS: # CURRENCY_COLUMNS should be defined in your class or globally
                    try:
                        return (2, float(str_value.replace('$', '').replace(',', ''))) # Type 2 for numbers
                    except ValueError:
                        return (3, str_value.lower()) # Type 3 for strings (malformed currency)

                # Attempt general numeric conversion for other columns
                try:
                    return (2, float(str_value)) # Type 2 for numbers
                except ValueError:
                    return (3, str_value.lower()) # Type 3 for strings (default)

            items_to_sort.sort(key=sort_key, reverse=reverse)

            # Reorder items in the treeview
            for index, (val, item_id) in enumerate(items_to_sort):
                self.tree.move(item_id, '', index)

            # Update the heading command to toggle sort direction
            self.tree.heading(col, command=lambda _col=col: self.sort_treeview_column(_col, not reverse))

        except Exception as e:
            logging.error(f"Error sorting column {col}: {e}", exc_info=True)
            # Optionally, reset the sort command on error to prevent repeated failures
            # self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))

    def center_toplevel(self, toplevel_window):
        """Centers a Toplevel window relative to the main application window."""
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
        """Handles the "Exit" menu command, prompting to save changes."""
        # Check if there are unsaved changes (conceptual - actual check might be more complex)
        # For simplicity, always ask, or you could implement a 'dirty' flag.
        if messagebox.askyesnocancel("Quit", "Do you want to save changes before quitting?"):
            self.save_data() # Save if "Yes"
            self.destroy()
        elif _ == False: # If "No" (askyesnocancel returns True for Yes, False for No, None for Cancel)
            self.destroy()
        # If None (Cancel), do nothing.

#for testing purposes only 
if __name__ == '__main__':
    app = OpenJobsApp()
    app.mainloop()