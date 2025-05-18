# data_management_tab.py
import tkinter as tk
from tkinter import ttk, messagebox, StringVar
import pandas as pd
import logging

import config # For accessing configurations

class DataManagementTab(ttk.Frame):
    """
    Manages the UI and interactions for the Data Management tab,
    including the Treeview display of job data.
    """
    def __init__(self, parent_notebook, app_instance):
        """
        Initialize the DataManagementTab.
        Args:
            parent_notebook: The ttk.Notebook widget this tab will belong to.
            app_instance: The instance of the main application (OpenJobsApp/AppShell).
        """
        super().__init__(parent_notebook)
        self.app = app_instance  # Store a reference to the main app
        self.tree = None         # Treeview widget will be initialized in _setup_ui
        self.editing_window = None # To manage the pop-up editor window

        self._setup_ui()

    def _setup_ui(self):
        """Creates and configures the widgets for this tab."""
        # Treeview and scrollbars
        self.tree = ttk.Treeview(self, columns=config.EXPECTED_COLUMNS, show="headings")
        for col in config.EXPECTED_COLUMNS:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            self.tree.column(col, width=config.PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout within this tab's frame
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        self.tree.grid(row=0, column=0, sticky='nsew', padx=config.DEFAULT_PADDING, pady=config.DEFAULT_PADDING)
        vsb.grid(row=0, column=1, sticky='ns', pady=config.DEFAULT_PADDING)
        hsb.grid(row=1, column=0, sticky='ew', padx=config.DEFAULT_PADDING)

        # Bind events
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.handle_delete_key) 

    def configure_treeview_columns(self):
        """Ensures the Treeview columns match config.EXPECTED_COLUMNS."""
        if not self.tree: return
        current_tree_cols = list(config.EXPECTED_COLUMNS)
        self.tree.configure(columns=current_tree_cols)
        for col in current_tree_cols:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            self.tree.column(col, width=config.PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)
        if '#0' in self.tree.column('#0'): 
            self.tree.column('#0', width=0, stretch=tk.NO)

    def populate_treeview(self):
        """Clears and repopulates the Treeview with data from self.app.status_df."""
        if not self.tree: 
            logging.warning("DMT: populate_treeview called but tree is not initialized.")
            return
        for i in self.tree.get_children():
            self.tree.delete(i)

        if self.app.status_df is None or self.app.status_df.empty:
            logging.info("DataManagementTab: No data to populate in the treeview.")
            return

        display_df = self.app.status_df.reindex(columns=config.EXPECTED_COLUMNS)
        date_columns = ['Order Date', 'Turn in Date'] 

        for df_index, row in display_df.iterrows(): 
            values = []
            for col_name in config.EXPECTED_COLUMNS:
                value = row[col_name]
                if col_name in date_columns and pd.notna(value):
                    try: value = pd.to_datetime(value).strftime(config.DATE_FORMAT)
                    except (ValueError, TypeError): logging.warning(f"DMT: Could not format date for '{value}' in col '{col_name}'. Original used.")
                elif col_name in config.CURRENCY_COLUMNS and pd.notna(value):
                    try:
                        # Corrected: remove regex=False for standard string replace
                        num_value = float(str(value).replace('$', '').replace(',', ''))
                        value = config.CURRENCY_FORMAT.format(num_value)
                    except (ValueError, TypeError): logging.warning(f"DMT: Could not format currency for '{value}' in col '{col_name}'. Original used.")
                values.append(value if pd.notna(value) else "")
            self.tree.insert("", tk.END, values=tuple(values), tags=(str(df_index),)) 

        self.app.after(10, self.set_column_widths_from_preferred) 
        self.color_rows()

    def on_double_click(self, event):
        if not self.tree: return
        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)
        if not item_id or not column_id_str: return

        try:
            column_index_tree = int(column_id_str.replace("#", "")) - 1
            if not (0 <= column_index_tree < len(config.EXPECTED_COLUMNS)):
                logging.warning(f"DMT: Invalid column index from tree: {column_index_tree}")
                return
            actual_column_name = config.EXPECTED_COLUMNS[column_index_tree]
            
            tags = self.tree.item(item_id, "tags")
            if not tags:
                logging.error(f"DMT: Item {item_id} has no tags (df_index missing). Cannot edit.")
                return
            df_row_index = int(tags[0]) 

            if self.editing_window and self.editing_window.winfo_exists():
                self.editing_window.destroy()
            self.editing_window = None

            if actual_column_name == "Status":
                self.create_status_editor(item_id, df_row_index, actual_column_name)
            elif actual_column_name == "Notes":
                self.create_notes_editor(item_id, df_row_index, actual_column_name)
            else:
                logging.debug(f"DMT: No special editor for column '{actual_column_name}'.")
        except (ValueError, IndexError, TypeError) as e:
            logging.error(f"DMT: Error in on_double_click: {e}. ItemID: {item_id}, ColIDStr: {column_id_str}", exc_info=True)

    def _save_edited_data(self, item_id, df_row_index, column_name, new_value, editor_window):
        try:
            self.app.perform_data_update(df_row_index, column_name, new_value)
            
            current_tree_values = list(self.tree.item(item_id, "values"))
            column_tree_idx = config.EXPECTED_COLUMNS.index(column_name)
            
            display_value = new_value
            if column_name in config.CURRENCY_COLUMNS and pd.notna(new_value):
                try: 
                    # Corrected: remove regex=False for standard string replace
                    display_value = config.CURRENCY_FORMAT.format(float(str(new_value).replace('$', '').replace(',', '')))
                except: pass 
            elif column_name in ['Order Date', 'Turn in Date'] and pd.notna(new_value):
                try: display_value = pd.to_datetime(new_value).strftime(config.DATE_FORMAT)
                except: pass
            current_tree_values[column_tree_idx] = display_value if pd.notna(display_value) else ""
            self.tree.item(item_id, values=tuple(current_tree_values))

            editor_window.destroy()
            self.editing_window = None
            self.color_rows() 
            self.app.notify_data_changed() 
        except Exception as e:
            logging.error(f"DMT: Error in _save_edited_data for column '{column_name}': {e}", exc_info=True)
            messagebox.showerror("Save Error", f"Could not save change for {column_name}: {e}", parent=editor_window)


    def create_notes_editor(self, item_id, df_row_index, column_name):
        self.editing_window = tk.Toplevel(self.app) 
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self.app); self.editing_window.grab_set()

        text_widget = tk.Text(self.editing_window, width=60, height=10, wrap=tk.WORD, font=config.DEFAULT_FONT)
        if df_row_index not in self.app.status_df.index:
            messagebox.showerror("Error", f"Invalid data index {df_row_index} for notes editor.", parent=self.app)
            self.editing_window.destroy()
            return
        current_value = str(self.app.status_df.loc[df_row_index, column_name])
        text_widget.insert(tk.END, current_value if pd.notna(current_value) else "")
        text_widget.pack(padx=config.DEFAULT_PADDING, pady=config.DEFAULT_PADDING, fill=tk.BOTH, expand=True)
        text_widget.focus()

        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(0,config.DEFAULT_PADDING), padx=config.DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._save_edited_data(item_id, df_row_index, column_name,
                                                                      text_widget.get("1.0", tk.END).strip(),
                                                                      self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.app.center_toplevel(self.editing_window) 

    def create_status_editor(self, item_id, df_row_index, column_name):
        self.editing_window = tk.Toplevel(self.app)
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self.app); self.editing_window.grab_set()

        if df_row_index not in self.app.status_df.index:
            messagebox.showerror("Error", f"Invalid data index {df_row_index} for status editor.", parent=self.app)
            self.editing_window.destroy()
            return
        current_value = str(self.app.status_df.loc[df_row_index, column_name])
        status_var = StringVar(self.editing_window)
        if current_value in config.ALLOWED_STATUS: status_var.set(current_value)
        elif config.ALLOWED_STATUS: status_var.set(config.ALLOWED_STATUS[0])
        else: status_var.set("")

        inv_num = self.app.status_df.loc[df_row_index, 'Invoice #'] 
        ttk.Label(self.editing_window, text=f"Status for Invoice {inv_num}:").pack(padx=config.DEFAULT_PADDING,pady=(config.DEFAULT_PADDING,5))
        combo = ttk.Combobox(self.editing_window, textvariable=status_var, values=config.ALLOWED_STATUS, state="readonly", font=config.DEFAULT_FONT)
        combo.pack(padx=config.DEFAULT_PADDING, pady=5, fill=tk.X); combo.focus()

        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(5,config.DEFAULT_PADDING), padx=config.DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._save_edited_data(item_id, df_row_index, column_name,
                                                                      status_var.get(), self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.app.center_toplevel(self.editing_window)

    def handle_delete_key(self, event=None):
        if not self.tree: return
        selected_tree_items = self.tree.selection()
        if not selected_tree_items:
            messagebox.showinfo("No Selection", "Please select one or more rows to delete.", parent=self.app)
            return

        confirm_msg = f"Are you sure you want to delete {len(selected_tree_items)} selected row(s)?"
        if not messagebox.askyesno("Confirm Delete", confirm_msg, parent=self.app):
            return

        df_indices_to_delete = []
        for item_id in selected_tree_items:
            tags = self.tree.item(item_id, "tags")
            if tags:
                try: df_indices_to_delete.append(int(tags[0]))
                except ValueError: logging.warning(f"DMT: Invalid tag for df_index on item {item_id}: {tags[0]}")
            else: logging.warning(f"DMT: Item {item_id} has no df_index tag. Cannot delete.")
        
        if not df_indices_to_delete:
            messagebox.showwarning("Deletion Error", "Could not identify valid rows to delete.", parent=self.app)
            return

        self.app.perform_delete_rows(df_indices_to_delete)
        self.populate_treeview() 
        messagebox.showinfo("Success", f"{len(df_indices_to_delete)} row(s) deleted.", parent=self.app)
        self.app.notify_data_changed() 

    def set_column_widths_from_preferred(self):
        if not self.tree: return
        self.app.update_idletasks() 
        for col_name in config.EXPECTED_COLUMNS:
            if col_name == '#0': continue 
            width = config.PREFERRED_COLUMN_WIDTHS.get(col_name, 100) 
            final_width = max(config.MIN_COLUMN_WIDTH, int(width))
            final_width = min(config.MAX_COLUMN_WIDTH, int(final_width))
            self.tree.column(col_name, width=final_width, anchor=tk.W)
        if '#0' in self.tree['columns'] and self.tree.column('#0', 'width') != 0 :
             self.tree.column('#0', width=0, stretch=tk.NO)

    def color_rows(self):
        if not self.tree or self.app.status_df is None or self.app.status_df.empty: return

        styles_map = {
            "default_status_style": (config.STATUS_COLORS["default_bg"], config.STATUS_COLORS["default_fg"]),
            "action_needed_style": (config.STATUS_COLORS["action_needed_bg"], config.STATUS_COLORS["action_needed_fg"]),
            "all_good_style": (config.STATUS_COLORS["all_good_bg"], config.STATUS_COLORS["all_good_fg"]),
            "closed_style": (config.STATUS_COLORS["closed_bg"], config.STATUS_COLORS["closed_fg"]),
            "new_style": (config.STATUS_COLORS["new_bg"], config.STATUS_COLORS["new_fg"]),
            "review_missing_style": (config.STATUS_COLORS["review_missing_bg"], config.STATUS_COLORS["review_missing_fg"])
        }
        for tag_name, (bg, fg) in styles_map.items():
            self.tree.tag_configure(tag_name, background=bg, foreground=fg)

        action_statuses = ["Ready to order", "Permit", "Waiting Measure"] 
        good_statuses = ["Ready to dispatch", "In install", "Done", "Waiting for materials"]
        status_col_name = "Status" 
        
        try: status_column_tree_index = config.EXPECTED_COLUMNS.index(status_col_name)
        except ValueError:
            logging.error(f"DMT: '{status_col_name}' column not found in EXPECTED_COLUMNS. Cannot color rows.")
            return

        for item_id in self.tree.get_children():
            try:
                df_index_tag_list = self.tree.item(item_id, "tags")
                df_index_tag = df_index_tag_list[0] if df_index_tag_list else f"err_idx_{item_id}"
                new_tags_for_item = [df_index_tag] 

                values = self.tree.item(item_id, "values")
                if values and len(values) > status_column_tree_index:
                    status = str(values[status_column_tree_index])
                    if status == "New": new_tags_for_item.append("new_style")
                    elif status == "Closed" or status == "Cancelled/Postponed": new_tags_for_item.append("closed_style")
                    elif status == config.REVIEW_MISSING_STATUS: new_tags_for_item.append("review_missing_style")
                    elif status in action_statuses: new_tags_for_item.append("action_needed_style")
                    elif status in good_statuses: new_tags_for_item.append("all_good_style")
                    else: new_tags_for_item.append("default_status_style")
                else:
                    new_tags_for_item.append("default_status_style")
                self.tree.item(item_id, tags=tuple(new_tags_for_item))
            except Exception as e:
                logging.error(f"DMT: Error coloring row {item_id}: {e}", exc_info=True)
                df_idx_tag_fallback = self.tree.item(item_id,"tags")[0] if self.tree.item(item_id,"tags") else f"err_idx_fb_{item_id}"
                self.tree.item(item_id, tags=(df_idx_tag_fallback, "default_status_style"))

    def sort_treeview_column(self, col, reverse):
        if not self.tree or self.app.status_df is None or self.app.status_df.empty: return

        date_columns = ['Order Date', 'Turn in Date'] 

        try:
            items_to_sort = []
            for item_id in self.tree.get_children(''):
                sort_value = None
                tags = self.tree.item(item_id, "tags")
                if not tags: 
                    logging.warning(f"DMT: Cannot sort item {item_id}, missing df_index tag. Using displayed value.")
                    sort_value = self.tree.set(item_id, col) 
                else:
                    try:
                        df_index = int(tags[0])
                        if df_index not in self.app.status_df.index: 
                             logging.warning(f"DMT: Invalid df_index {df_index} for item {item_id} during sort. Using displayed value.")
                             sort_value = self.tree.set(item_id, col)
                        else:
                            original_value = self.app.status_df.loc[df_index, col]
                            if col in date_columns:
                                sort_value = pd.to_datetime(original_value, errors='coerce')
                            elif col in config.CURRENCY_COLUMNS:
                                try: 
                                    # Corrected: remove regex=False for standard string replace
                                    sort_value = float(str(original_value).replace('$', '').replace(',', ''))
                                except (ValueError, TypeError): sort_value = str(original_value) 
                            else: 
                                try: sort_value = float(original_value)
                                except (ValueError, TypeError): sort_value = str(original_value)
                    except (IndexError, KeyError, ValueError, TypeError) as e:
                        logging.warning(f"DMT: Sort fallback for col {col}, item {item_id} (df_idx {tags[0]}): {e}. Using displayed value.")
                        sort_value = self.tree.set(item_id, col) 
                
                items_to_sort.append((sort_value, item_id))

            def sort_key(item_tuple):
                value = item_tuple[0]
                if isinstance(value, pd.Timestamp): return (0, value) 
                if pd.isna(value): return (1, pd.Timestamp.min if not reverse else pd.Timestamp.max) 
                if isinstance(value, (int, float)): return (2, value) 
                return (3, str(value).lower()) 

            items_to_sort.sort(key=sort_key, reverse=reverse)
            for index, (val, item_id) in enumerate(items_to_sort):
                self.tree.move(item_id, '', index)
            self.tree.heading(col, command=lambda _col=col: self.sort_treeview_column(_col, not reverse))
        except Exception as e:
            logging.error(f"DMT: Error sorting column {col}: {e}", exc_info=True)

    def on_tab_selected(self):
        logging.info("Data Management tab selected.")
        if self.tree: 
            if not self.tree.get_children() and self.app.status_df is not None and not self.app.status_df.empty:
                logging.info("Data Management tab was empty, populating treeview on selection.")
                self.populate_treeview()
            else: 
                self.configure_treeview_columns()
                self.set_column_widths_from_preferred()
        else:
            logging.warning("Data Management tab selected, but treeview not initialized.")

