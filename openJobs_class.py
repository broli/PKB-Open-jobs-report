import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar, OptionMenu
import sv_ttk # Import the sun valley ttk theme
import pandas as pd
from data_utils import load_status, save_status, load_excel, process_data
import logging

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
OUTPUT_FILE = "open_invoices.xlsx" # Default name for generated report
EXPECTED_COLUMNS = [
    '#', 'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]
DEFAULT_FONT = ('Calibri', 12) # Adjusted default font size for potentially more columns visible
DEFAULT_PADDING = 10
DATE_FORMAT = '%b-%d'
CURRENCY_COLUMNS = ['Invoice Total', 'Balance']
CURRENCY_FORMAT = '${:,.2f}'
ALLOWED_STATUS = [
    "Waiting Measure", "Ready to order", "Waiting for materials",
    "Ready to dispatch", "In install", "Done", "Permit",
    "Cancelled/Postponed", "New", "Closed"
]

# Define preferred widths for specific columns (in pixels)
# Adjust these values to your liking.
PREFERRED_COLUMN_WIDTHS = {
    '#': 10,
    'Invoice #': 30,
    'Order Date': 30,
    'Turn in Date': 30,
    'Account': 220,
    'Invoice Total': 50,
    'Balance': 50,
    'Salesperson': 100,
    'Project Coordinator': 130,
    'Status': 160,
    'Notes': 300
}

# Minimum and maximum allowable width for any column (dynamic or preferred)
MIN_COLUMN_WIDTH = 10
MAX_COLUMN_WIDTH = 450 # Adjusted max slightly

# Configure the logging level
logging.basicConfig(level=logging.DEBUG)

class OpenJobsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        sv_ttk.set_theme("light")
        self.title("Open Jobs App")
        self.maximize_window()

        self.style = ttk.Style(self)
        self.style.configure("Treeview", font=DEFAULT_FONT, rowheight=28) # Adjusted rowheight
        self.style.configure("Treeview.Heading", font=(DEFAULT_FONT[0], DEFAULT_FONT[1], 'bold'))

        self.status_colors = {
            "default_fg": "black", "default_bg": "white",
            "action_needed_bg": "#FFEBEE", "action_needed_fg": "black",
            "all_good_bg": "#E8F5E9", "all_good_fg": "black",
            "closed_bg": "#F5F5F5", "closed_fg": "black",
            "new_bg": "#E3F2FD", "new_fg": "black",
            "selected_bg": "#B0BEC5", "selected_fg": "black"
        }
        
        self.style.map("Treeview",
                       background=[('selected', self.status_colors["selected_bg"])],
                       foreground=[('selected', self.status_colors["selected_fg"])])

        self.status_df = load_status()
        if self.status_df is None:
            messagebox.showerror("Fatal Error", "Could not load status data. Exiting.")
            self.destroy()
            return
        elif not all(col in self.status_df.columns for col in EXPECTED_COLUMNS):
             messagebox.showwarning("Data Warning", "Loaded data is missing some expected columns. Defaults will be used.")
             for col in EXPECTED_COLUMNS:
                 if col not in self.status_df.columns:
                     self.status_df[col] = "" # Use empty string as default for missing cols

        self.create_widgets()

    def maximize_window(self):
        try:
            self.state('zoomed')
        except tk.TclError:
            self.attributes('-fullscreen', True)

    def create_widgets(self):
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel)
        file_menu.add_command(label="Save", command=self.save_data)
        file_menu.add_command(label="Generate Report", command=self.generate_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menu_bar)

        tree_columns = list(self.status_df.columns) if self.status_df is not None else EXPECTED_COLUMNS
        self.tree = ttk.Treeview(self, columns=tree_columns, show="headings")

        for col in tree_columns:
            self.tree.heading(col, text=col)
            # Initial width is set here but will be overridden by set_column_widths
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected_row)
        self.editing_window = None

        self.populate_treeview()

    def configure_treeview_columns(self):
        current_columns = list(self.status_df.columns)
        self.tree.configure(columns=current_columns)
        for col in current_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)

    def populate_treeview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if self.status_df is None or self.status_df.empty:
            logging.info("No data to populate in the treeview.")
            self.set_column_widths() # Still set widths for empty table based on headings/preferred
            return

        date_columns = ['Order Date', 'Turn in Date']
        try:
            status_col_idx = list(self.status_df.columns).index('Status')
        except ValueError:
            logging.error("'Status' column not found in DataFrame. Cannot populate correctly.")
            return

        for index, row in self.status_df.iterrows():
            values = []
            for col_name in self.status_df.columns:
                value = row[col_name]
                if col_name in date_columns and pd.notna(value):
                    try: value = pd.to_datetime(value).strftime(DATE_FORMAT)
                    except: pass
                elif col_name in CURRENCY_COLUMNS and pd.notna(value):
                    try:
                        value = float(str(value).replace('$', '').replace(',', ''))
                        value = CURRENCY_FORMAT.format(value)
                    except: pass
                values.append(value if pd.notna(value) else "")
            self.tree.insert("", tk.END, values=tuple(values), tags=(str(index),))

        self.after(10, self.set_column_widths)
        self.color_rows()

    def load_new_excel(self):
        excel_file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if not excel_file_path: return

        new_df = load_excel(excel_file_path)
        if new_df is None: return

        self.status_df = process_data(new_df, self.status_df.copy() if self.status_df is not None else pd.DataFrame(columns=EXPECTED_COLUMNS))

        if self.status_df is None:
            messagebox.showerror("Error", "Failed to process the new Excel data.")
            self.status_df = pd.DataFrame(columns=EXPECTED_COLUMNS)
            return

        for col in EXPECTED_COLUMNS:
            if col not in self.status_df.columns:
                self.status_df[col] = "" # Default for new columns

        self.status_df = self.status_df[EXPECTED_COLUMNS] # Ensure correct column order
        self.configure_treeview_columns()
        self.populate_treeview()

    def save_data(self):
        if self.status_df is None:
            messagebox.showerror("Error", "No data to save.")
            return
        save_status(self.status_df)

    def generate_report(self):
        if self.status_df is None or self.status_df.empty:
            messagebox.showinfo("Info", "No data available to generate a report.")
            return

        open_invoices = self.status_df[
            ~self.status_df['Status'].isin(['Closed', 'Cancelled/Postponed'])
        ].copy()

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
            open_invoices.to_excel(report_file_path, index=False)
            messagebox.showinfo("Success", f"Report generated: {report_file_path}")
        except Exception as e:
            logging.error(f"Error generating report: {e}")
            messagebox.showerror("Error", f"Error generating report: {e}")

    def on_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)
        if not item_id or not column_id_str: return

        try:
            column_index = int(column_id_str.replace("#", "")) - 1
            if not (0 <= column_index < len(self.status_df.columns)): return

            actual_column_name = self.status_df.columns[column_index]
            df_row_index = int(self.tree.item(item_id, "tags")[0])

            if self.editing_window and self.editing_window.winfo_exists():
                self.editing_window.destroy()
            self.editing_window = None

            if actual_column_name == "Status":
                self.create_status_dropdown(item_id, df_row_index, column_index, actual_column_name)
            elif actual_column_name == "Notes":
                self.create_notes_editor(item_id, df_row_index, column_index, actual_column_name)
            else:
                logging.debug(f"No special editor for column '{actual_column_name}'.")
        except (ValueError, IndexError) as e:
            logging.error(f"Error in on_double_click: {e}")

    def _common_editor_save(self, item_id, df_row_index, column_idx, new_value, editor_window):
        """Helper to save edit from editor windows."""
        self.status_df.iloc[df_row_index, column_idx] = new_value
        current_tree_values = list(self.tree.item(item_id, "values"))
        current_tree_values[column_idx] = new_value
        self.tree.item(item_id, values=tuple(current_tree_values))
        editor_window.destroy()
        self.editing_window = None
        self.color_rows()

    def create_notes_editor(self, item_id, df_row_index, column_idx, actual_column_name):
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {actual_column_name}")
        self.editing_window.transient(self); self.editing_window.grab_set()

        current_value = str(self.status_df.iloc[df_row_index, column_idx])
        text_widget = tk.Text(self.editing_window, width=60, height=10, font=('Calibri', 12), wrap=tk.WORD)
        text_widget.insert(tk.END, current_value); text_widget.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING, fill=tk.BOTH, expand=True); text_widget.focus()
        
        btn_frame = ttk.Frame(self.editing_window); btn_frame.pack(pady=(0,DEFAULT_PADDING),padx=DEFAULT_PADDING,fill=tk.X)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton", command=lambda: self._common_editor_save(item_id, df_row_index, column_idx, text_widget.get("1.0", tk.END).strip(), self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None))); cancel_btn.pack(side=tk.RIGHT)
        self.editing_window.protocol("WM_DELETE_WINDOW", lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.wait_window(self.editing_window)

    def create_status_dropdown(self, item_id, df_row_index, column_index, column_name):
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {column_name}"); self.editing_window.transient(self); self.editing_window.grab_set()

        current_value = str(self.status_df.iloc[df_row_index, column_index])
        status_var = StringVar(self.editing_window)
        status_var.set(current_value if current_value in ALLOWED_STATUS else (ALLOWED_STATUS[0] if ALLOWED_STATUS else ""))
        
        inv_num = self.status_df.iloc[df_row_index].get('Invoice #', 'N/A')
        ttk.Label(self.editing_window, text=f"Status for Invoice {inv_num}:").pack(padx=DEFAULT_PADDING,pady=(DEFAULT_PADDING,5))
        
        combo = ttk.Combobox(self.editing_window, textvariable=status_var, values=ALLOWED_STATUS, state="readonly", font=DEFAULT_FONT)
        combo.pack(padx=DEFAULT_PADDING, pady=5, fill=tk.X); combo.focus()
        
        btn_frame = ttk.Frame(self.editing_window); btn_frame.pack(pady=(5,DEFAULT_PADDING),padx=DEFAULT_PADDING,fill=tk.X)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton", command=lambda: self._common_editor_save(item_id, df_row_index, column_index, status_var.get(), self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None))); cancel_btn.pack(side=tk.RIGHT)
        self.editing_window.protocol("WM_DELETE_WINDOW", lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.wait_window(self.editing_window)

    def delete_selected_row(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items: messagebox.showinfo("No Selection", "Please select row(s) to delete."); return
        if not messagebox.askyesno("Confirm Delete", f"Delete {len(selected_items)} row(s)? This cannot be undone from UI."): return

        df_indices_to_delete = sorted([int(self.tree.item(item_id, "tags")[0]) for item_id in selected_items], reverse=True)
        
        for item_id in selected_items: self.tree.delete(item_id)
        if df_indices_to_delete:
            self.status_df.drop(index=df_indices_to_delete, inplace=True)
            self.status_df.reset_index(drop=True, inplace=True)
        
        self.re_tag_tree_items(); self.color_rows()

    def re_tag_tree_items(self):
        for new_df_index, item_id in enumerate(self.tree.get_children()):
            current_tags = list(self.tree.item(item_id, 'tags'))
            updated_tags = [tag for tag in current_tags if not tag.isdigit()] # Remove old numeric tags
            updated_tags.insert(0, str(new_df_index)) # Add new index tag
            self.tree.item(item_id, tags=tuple(updated_tags))

    def set_column_widths(self):
        self.update_idletasks()
        if self.status_df is None: return # Guard clause

        for col_name in self.tree['columns']: # Iterate over actual treeview columns
            if col_name == '#0': continue # Skip the special #0 column

            final_width = PREFERRED_COLUMN_WIDTHS.get(col_name)

            if final_width is None: # Dynamic calculation
                try:
                    heading_text = self.tree.heading(col_name, "text")
                    temp_label_heading = ttk.Label(self, text=heading_text, font=(DEFAULT_FONT[0], DEFAULT_FONT[1], 'bold'))
                    heading_width = temp_label_heading.winfo_reqwidth() + 15
                    temp_label_heading.destroy()

                    max_content_width = 0
                    # Ensure col_name is valid for status_df if dynamically accessing by name
                    if col_name in self.status_df.columns:
                        col_idx = list(self.status_df.columns).index(col_name)
                        items_to_check = self.tree.get_children()
                        # Sample for performance if many items
                        if len(items_to_check) > 50: items_to_check = items_to_check[:30] + items_to_check[-20:]

                        for item_id in items_to_check:
                            values = self.tree.item(item_id, "values")
                            if values and col_idx < len(values): # Check col_idx bounds
                                value_text = str(values[col_idx])
                                temp_label_val = ttk.Label(self, text=value_text, font=DEFAULT_FONT)
                                content_width = temp_label_val.winfo_reqwidth()
                                temp_label_val.destroy()
                                if content_width > max_content_width: max_content_width = content_width
                        final_width = max(heading_width, max_content_width + 25)
                    else: # Column in tree but not in df (should not happen with configure_treeview_columns)
                        final_width = heading_width # Base on heading if data column is missing
                except Exception as e:
                    logging.warning(f"Dynamic width calc error for '{col_name}': {e}")
                    final_width = 100 # Fallback
            
            final_width = max(MIN_COLUMN_WIDTH, int(final_width))
            final_width = min(MAX_COLUMN_WIDTH, int(final_width))
            self.tree.column(col_name, width=final_width, anchor=tk.W)
        self.tree.column('#0', width=0, stretch=tk.NO)

    def color_rows(self):
        if self.status_df is None: return

        styles = {"default_status_style": (self.status_colors["default_bg"], self.status_colors["default_fg"]),
                  "action_needed_style": (self.status_colors["action_needed_bg"], self.status_colors["action_needed_fg"]),
                  "all_good_style": (self.status_colors["all_good_bg"], self.status_colors["all_good_fg"]),
                  "closed_style": (self.status_colors["closed_bg"], self.status_colors["closed_fg"]),
                  "new_style": (self.status_colors["new_bg"], self.status_colors["new_fg"])}
        for tag_name, (bg, fg) in styles.items(): self.tree.tag_configure(tag_name, background=bg, foreground=fg)

        default_statuses = ["Waiting Measure", "Waiting for materials"]
        action_statuses = ["Ready to order", "Permit", "Cancelled/Postponed"]
        good_statuses = ["Ready to dispatch", "In install", "Done"]
        
        try: status_column_index = list(self.status_df.columns).index("Status")
        except ValueError: logging.error("'Status' column missing. Cannot color rows."); return

        for item_id in self.tree.get_children():
            try:
                df_index_tag = self.tree.item(item_id, "tags")[0] # Assumes index tag is first
                new_tags = [df_index_tag]
                values = self.tree.item(item_id, "values")

                if values and len(values) > status_column_index:
                    status = str(values[status_column_index])
                    if status in default_statuses: new_tags.append("default_status_style")
                    elif status in action_statuses: new_tags.append("action_needed_style")
                    elif status in good_statuses: new_tags.append("all_good_style")
                    elif status == "Closed": new_tags.append("closed_style")
                    elif status == "New": new_tags.append("new_style")
                    else: new_tags.append("default_status_style")
                else: new_tags.append("default_status_style")
                self.tree.item(item_id, tags=tuple(new_tags))
            except Exception as e:
                logging.error(f"Error coloring row {item_id}: {e}")
                df_idx_tag = self.tree.item(item_id,"tags")[0] if self.tree.item(item_id,"tags") else "err_idx"
                self.tree.item(item_id, tags=(df_idx_tag, "default_status_style"))


if __name__ == '__main__':
    # This part is for testing the class directly,
    # your Main.py will be the primary way to run the app.
    app = OpenJobsApp()
    app.mainloop()