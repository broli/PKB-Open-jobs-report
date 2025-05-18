# data_management_tab.py
# This module handles everything related to the 'Data Management' tab in the application.
# It's responsible for displaying the job data in a treeview, allowing users to
# interact with it (edit, delete, sort), and ensuring the data presentation is
# clear and user-friendly.

import tkinter as tk
from tkinter import ttk, messagebox, StringVar # Ensure StringVar is imported for dynamic UI text
import pandas as pd
import logging

import config # For accessing configurations like column names, colors, etc.

class DataManagementTab(ttk.Frame):
    """
    Manages the UI and interactions for the Data Management tab.
    This includes the Treeview display of job data, editing popups,
    and row styling based on status.
    """
    def __init__(self, parent_notebook, app_instance):
        """
        Initialize the DataManagementTab.
        Args:
            parent_notebook: The ttk.Notebook widget this tab will belong to.
            app_instance: The instance of the main application (OpenJobsApp/AppShell),
                          used to access shared data and methods.
        """
        super().__init__(parent_notebook)
        self.app = app_instance  # Store a reference to the main app
        self.tree = None         # Treeview widget will be initialized in _setup_ui
        self.editing_window = None # To manage the pop-up editor window, ensuring only one is open at a time

        self._setup_ui() # Build the user interface for this tab

    def _setup_ui(self):
        """Creates and configures the widgets for this tab."""
        # Main Treeview widget to display the job data
        self.tree = ttk.Treeview(self, columns=config.EXPECTED_COLUMNS, show="headings")
        
        # Configure each column in the Treeview
        for col in config.EXPECTED_COLUMNS:
            # Set column heading text and enable sorting when a heading is clicked
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            # Set default width and alignment for each column
            self.tree.column(col, width=config.PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)

        # Scrollbars for the Treeview (vertical and horizontal)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout: Make the Treeview expand with the window
        self.columnconfigure(0, weight=1) # Treeview column
        self.rowconfigure(0, weight=1)    # Treeview row
        
        # Place Treeview and scrollbars in the grid
        self.tree.grid(row=0, column=0, sticky='nsew', padx=config.DEFAULT_PADDING, pady=config.DEFAULT_PADDING)
        vsb.grid(row=0, column=1, sticky='ns', pady=config.DEFAULT_PADDING) # Vertical scrollbar to the right
        hsb.grid(row=1, column=0, sticky='ew', padx=config.DEFAULT_PADDING) # Horizontal scrollbar below

        # Bind events to Treeview actions
        self.tree.bind("<Double-1>", self.on_double_click) # Double-click to edit a cell
        self.tree.bind("<Delete>", self.handle_delete_key) # Delete key to remove selected row(s)

    def configure_treeview_columns(self):
        """
        Ensures the Treeview columns match the current config.EXPECTED_COLUMNS.
        This can be useful if the expected columns change dynamically or need refreshing.
        """
        if not self.tree: return # Safety check
        current_tree_cols = list(config.EXPECTED_COLUMNS)
        self.tree.configure(columns=current_tree_cols) # Update the columns definition
        for col in current_tree_cols:
            # Re-apply heading text and sort command
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview_column(_col, False))
            # Re-apply column width and anchor
            self.tree.column(col, width=config.PREFERRED_COLUMN_WIDTHS.get(col, 100), anchor=tk.W)
        
        # Hide the default first column ('#0') if it's present and not already hidden
        if '#0' in self.tree.column('#0'): 
            self.tree.column('#0', width=0, stretch=tk.NO)

    def populate_treeview(self):
        """
        Clears and repopulates the Treeview with data from self.app.status_df.
        This is the main method to refresh the displayed data.
        """
        if not self.tree: 
            logging.warning("DMT: populate_treeview called but tree is not initialized.")
            return
        
        # Clear existing items from the Treeview
        for i in self.tree.get_children():
            self.tree.delete(i)

        # If no data is loaded in the app, nothing to show
        if self.app.status_df is None or self.app.status_df.empty:
            logging.info("DataManagementTab: No data to populate in the treeview.")
            return

        # Ensure the DataFrame columns are in the expected order for display
        display_df = self.app.status_df.reindex(columns=config.EXPECTED_COLUMNS)
        date_columns = ['Order Date', 'Turn in Date'] # Columns that need date formatting

        # Iterate through each row in the DataFrame and insert it into the Treeview
        for df_index, row in display_df.iterrows(): 
            values = [] # List to hold formatted values for the current row
            for col_name in config.EXPECTED_COLUMNS:
                value = row[col_name]
                # Apply specific formatting for date and currency columns
                if col_name in date_columns and pd.notna(value):
                    try: value = pd.to_datetime(value).strftime(config.DATE_FORMAT)
                    except (ValueError, TypeError): logging.warning(f"DMT: Could not format date for '{value}' in col '{col_name}'. Original used.")
                elif col_name in config.CURRENCY_COLUMNS and pd.notna(value):
                    try:
                        # Convert to float then format as currency
                        num_value = float(str(value).replace('$', '').replace(',', ''))
                        value = config.CURRENCY_FORMAT.format(num_value)
                    except (ValueError, TypeError): logging.warning(f"DMT: Could not format currency for '{value}' in col '{col_name}'. Original used.")
                
                # Use empty string for NaN/None values in display
                values.append(value if pd.notna(value) else "")
            
            # Insert the row into the Treeview. The original DataFrame index is stored as a tag.
            self.tree.insert("", tk.END, values=tuple(values), tags=(str(df_index),)) 

        # Adjust column widths after data is populated (deferred for accurate calculation)
        self.app.after(10, self.set_column_widths_from_preferred) 
        self.color_rows() # Apply row styling based on status

    def on_double_click(self, event):
        """
        Handles a double-click event on a Treeview cell.
        Identifies the clicked row and column, then opens an appropriate editor.
        """
        if not self.tree: return # Safety check

        # Identify the row and column clicked
        item_id = self.tree.identify_row(event.y) # Treeview item identifier
        column_id_str = self.tree.identify_column(event.x) # Column identifier (e.g., "#1")
        
        if not item_id or not column_id_str: return # Click was not on a valid cell

        try:
            # Convert column identifier to an index and get the column name
            column_index_tree = int(column_id_str.replace("#", "")) - 1 # Treeview columns are 1-indexed
            if not (0 <= column_index_tree < len(config.EXPECTED_COLUMNS)):
                logging.warning(f"DMT: Invalid column index from tree: {column_index_tree}")
                return
            actual_column_name = config.EXPECTED_COLUMNS[column_index_tree]
            
            # Retrieve the original DataFrame index stored in the item's tags
            tags = self.tree.item(item_id, "tags")
            if not tags:
                logging.error(f"DMT: Item {item_id} has no tags (df_index missing). Cannot edit.")
                return
            df_row_index = int(tags[0]) # The first tag is assumed to be the DataFrame index

            # Close any existing editor window before opening a new one
            if self.editing_window and self.editing_window.winfo_exists():
                self.editing_window.destroy()
            self.editing_window = None

            # Call the appropriate editor based on the column name
            if actual_column_name == "Status":
                self.create_status_editor(item_id, df_row_index, actual_column_name)
            elif actual_column_name == "Notes":
                self.create_notes_editor(item_id, df_row_index, actual_column_name)
            elif actual_column_name == "Project Coordinator": 
                self.create_pc_editor(item_id, df_row_index, actual_column_name)
            else:
                # For other columns, no special editor is defined (they are not user-editable this way)
                logging.debug(f"DMT: No special editor for column '{actual_column_name}'.")
        except (ValueError, IndexError, TypeError) as e:
            logging.error(f"DMT: Error in on_double_click: {e}. ItemID: {item_id}, ColIDStr: {column_id_str}", exc_info=True)

    def _save_edited_data(self, item_id, df_row_index, column_name, new_value, editor_window):
        """
        Saves the edited data back to the main DataFrame and updates the Treeview.
        This is a callback used by the editor popups.
        Args:
            item_id: The Treeview item identifier.
            df_row_index: The original DataFrame index for the row.
            column_name: The name of the column being edited.
            new_value: The new value to save.
            editor_window: The Toplevel editor window to be closed on save.
        """
        try:
            # Update the underlying DataFrame in the main application
            self.app.perform_data_update(df_row_index, column_name, new_value)
            
            # Update the value displayed in the Treeview
            current_tree_values = list(self.tree.item(item_id, "values"))
            column_tree_idx = config.EXPECTED_COLUMNS.index(column_name) # Get Treeview column index
            
            # Format the value for display if necessary (currency, date)
            display_value = new_value 
            if column_name in config.CURRENCY_COLUMNS and pd.notna(new_value):
                try: 
                    display_value = config.CURRENCY_FORMAT.format(float(str(new_value).replace('$', '').replace(',', '')))
                except: pass # If formatting fails, use raw new_value
            elif column_name in ['Order Date', 'Turn in Date'] and pd.notna(new_value):
                try: display_value = pd.to_datetime(new_value).strftime(config.DATE_FORMAT)
                except: pass
            
            current_tree_values[column_tree_idx] = display_value if pd.notna(display_value) else ""
            self.tree.item(item_id, values=tuple(current_tree_values))

            editor_window.destroy() # Close the editor popup
            self.editing_window = None
            self.color_rows() # Re-apply row styling as status might have changed
            self.app.notify_data_changed() # Inform other parts of the app (like Reporting tab)
        except Exception as e:
            logging.error(f"DMT: Error in _save_edited_data for column '{column_name}': {e}", exc_info=True)
            messagebox.showerror("Save Error", f"Could not save change for {column_name}: {e}", parent=editor_window)


    def create_notes_editor(self, item_id, df_row_index, column_name):
        """Creates a popup editor for the 'Notes' column using a Text widget."""
        self.editing_window = tk.Toplevel(self.app) 
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self.app); self.editing_window.grab_set() # Modal behavior

        # Text widget for multi-line notes
        text_widget = tk.Text(self.editing_window, width=60, height=10, wrap=tk.WORD, font=config.DEFAULT_FONT)
        
        # Ensure the DataFrame index is valid before trying to access data
        if df_row_index not in self.app.status_df.index:
            messagebox.showerror("Error", f"Invalid data index {df_row_index} for notes editor.", parent=self.app)
            self.editing_window.destroy()
            return
        
        current_value = str(self.app.status_df.loc[df_row_index, column_name])
        text_widget.insert(tk.END, current_value if pd.notna(current_value) else "") # Populate with current notes
        text_widget.pack(padx=config.DEFAULT_PADDING, pady=config.DEFAULT_PADDING, fill=tk.BOTH, expand=True)
        text_widget.focus() # Set focus to the text widget

        # Frame for Save and Cancel buttons
        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(0,config.DEFAULT_PADDING), padx=config.DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._save_edited_data(item_id, df_row_index, column_name,
                                                                      text_widget.get("1.0", tk.END).strip(), # Get all text
                                                                      self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5,0))
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)
        
        # Handle window close button (the 'X')
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        self.app.center_toplevel(self.editing_window) # Center the popup

    def create_status_editor(self, item_id, df_row_index, column_name):
        """Creates a popup editor for the 'Status' column using a Combobox."""
        self.editing_window = tk.Toplevel(self.app)
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self.app); self.editing_window.grab_set()

        if df_row_index not in self.app.status_df.index:
            messagebox.showerror("Error", f"Invalid data index {df_row_index} for status editor.", parent=self.app)
            self.editing_window.destroy()
            return
            
        current_value = str(self.app.status_df.loc[df_row_index, column_name])
        status_var = StringVar(self.editing_window) # Tkinter variable for Combobox
        
        # Set current status in Combobox, or default if not in allowed list
        if current_value in config.ALLOWED_STATUS: status_var.set(current_value)
        elif config.ALLOWED_STATUS: status_var.set(config.ALLOWED_STATUS[0]) # Default to first allowed status
        else: status_var.set("") # Should not happen if data is clean

        # Display Invoice # for context
        inv_num = self.app.status_df.loc[df_row_index, 'Invoice #'] 
        ttk.Label(self.editing_window, text=f"Status for Invoice {inv_num}:").pack(padx=config.DEFAULT_PADDING,pady=(config.DEFAULT_PADDING,5))
        
        # Combobox with predefined status values
        combo = ttk.Combobox(self.editing_window, textvariable=status_var, values=config.ALLOWED_STATUS, state="readonly", font=config.DEFAULT_FONT)
        combo.pack(padx=config.DEFAULT_PADDING, pady=5, fill=tk.X); combo.focus()

        # Save and Cancel buttons
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

    def create_pc_editor(self, item_id, df_row_index, column_name):
        """Creates a popup editor for the 'Project Coordinator' column using a Combobox."""
        self.editing_window = tk.Toplevel(self.app)
        self.editing_window.title(f"Edit {column_name}")
        self.editing_window.transient(self.app)
        self.editing_window.grab_set() # Make it modal

        # Validate DataFrame index
        if df_row_index not in self.app.status_df.index:
            messagebox.showerror("Error", f"Invalid data index {df_row_index} for {column_name} editor.", parent=self.app)
            self.editing_window.destroy()
            return

        # Get the current Project Coordinator value
        current_value = str(self.app.status_df.loc[df_row_index, column_name])
        pc_var = StringVar(self.editing_window) # Tkinter variable for the Combobox

        # Populate the list of Project Coordinators dynamically from the DataFrame
        pc_list = []
        if 'Project Coordinator' in self.app.status_df.columns:
            # Get unique, non-null, non-empty-string PC names, then sort them alphabetically
            pc_names_series = self.app.status_df['Project Coordinator'].dropna().unique()
            pc_list = sorted([str(name) for name in pc_names_series if str(name).strip()]) # Filter out empty strings
            if not pc_list:
                # Fallback if no valid PC names are found after filtering
                pc_list = ["No Coordinators Found"] 
        else:
            # Fallback if the 'Project Coordinator' column itself is missing (should be rare with EXPECTED_COLUMNS)
            pc_list = ["Project Coordinator Column Missing"]

        # Set the current value in the Combobox
        if current_value in pc_list:
            pc_var.set(current_value)
        elif pc_list and pc_list[0] not in ["No Coordinators Found", "Project Coordinator Column Missing"]:
            # If current value isn't in the list (e.g., it was blank or an old/invalid name),
            # default to the first valid PC name in the list if available.
            pc_var.set(pc_list[0])
        else:
            # If the list is empty or contains only error/fallback messages, set Combobox to empty.
            pc_var.set("") 

        # Display the Invoice # for context, so the user knows which job they're editing
        inv_num = self.app.status_df.loc[df_row_index, 'Invoice #'] 
        ttk.Label(self.editing_window, text=f"{column_name} for Invoice {inv_num}:").pack(padx=config.DEFAULT_PADDING, pady=(config.DEFAULT_PADDING, 5))
        
        # Create the Combobox for PC selection (read-only to enforce selection from the list)
        combo = ttk.Combobox(self.editing_window, textvariable=pc_var, values=pc_list, state="readonly", font=config.DEFAULT_FONT)
        combo.pack(padx=config.DEFAULT_PADDING, pady=5, fill=tk.X) # Fill horizontally
        combo.focus() # Set focus to the combobox

        # Frame for Save and Cancel buttons
        btn_frame = ttk.Frame(self.editing_window)
        btn_frame.pack(pady=(5, config.DEFAULT_PADDING), padx=config.DEFAULT_PADDING, fill=tk.X, side=tk.BOTTOM)
        
        # Save button: calls _save_edited_data with the selected PC name
        save_btn = ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                              command=lambda: self._save_edited_data(item_id, df_row_index, column_name,
                                                                      pc_var.get(), self.editing_window))
        save_btn.pack(side=tk.RIGHT, padx=(5, 0)) # Align right
        
        # Cancel button: closes the popup without saving
        cancel_btn = ttk.Button(btn_frame, text="Cancel",
                                command=lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        cancel_btn.pack(side=tk.RIGHT)
        
        # Ensure the editing_window attribute is cleared if the window is closed via 'X'
        self.editing_window.protocol("WM_DELETE_WINDOW",
                                     lambda: (self.editing_window.destroy(), setattr(self, 'editing_window', None)))
        # Center the popup window on the screen relative to the main app
        self.app.center_toplevel(self.editing_window)

    def handle_delete_key(self, event=None):
        """Handles the Delete key press to remove selected rows from the Treeview and DataFrame."""
        if not self.tree: return
        selected_tree_items = self.tree.selection() # Get all selected items
        if not selected_tree_items:
            messagebox.showinfo("No Selection", "Please select one or more rows to delete.", parent=self.app)
            return

        # Confirm deletion with the user
        confirm_msg = f"Are you sure you want to delete {len(selected_tree_items)} selected row(s)?"
        if not messagebox.askyesno("Confirm Delete", confirm_msg, parent=self.app):
            return

        # Collect DataFrame indices of rows to be deleted
        df_indices_to_delete = []
        for item_id in selected_tree_items:
            tags = self.tree.item(item_id, "tags") # Get DataFrame index from tags
            if tags:
                try: df_indices_to_delete.append(int(tags[0]))
                except ValueError: logging.warning(f"DMT: Invalid tag for df_index on item {item_id}: {tags[0]}")
            else: logging.warning(f"DMT: Item {item_id} has no df_index tag. Cannot delete.")
        
        if not df_indices_to_delete:
            messagebox.showwarning("Deletion Error", "Could not identify valid rows to delete.", parent=self.app)
            return

        # Perform deletion in the main application's DataFrame
        self.app.perform_delete_rows(df_indices_to_delete)
        self.populate_treeview() # Refresh the Treeview to reflect deletions
        messagebox.showinfo("Success", f"{len(df_indices_to_delete)} row(s) deleted.", parent=self.app)
        self.app.notify_data_changed() # Notify other parts of the application

    def set_column_widths_from_preferred(self):
        """
        Adjusts Treeview column widths based on preferred settings in config.
        Ensures widths are within min/max bounds.
        Called after data population for better accuracy.
        """
        if not self.tree: return
        self.app.update_idletasks() # Ensure Tkinter has processed pending geometry changes
        for col_name in config.EXPECTED_COLUMNS:
            if col_name == '#0': continue # Skip the hidden default column
            
            # Get preferred width, apply min/max constraints
            width = config.PREFERRED_COLUMN_WIDTHS.get(col_name, 100) # Default if not specified
            final_width = max(config.MIN_COLUMN_WIDTH, int(width))
            final_width = min(config.MAX_COLUMN_WIDTH, int(final_width))
            self.tree.column(col_name, width=final_width, anchor=tk.W)
            
        # Explicitly hide column #0 if it exists and isn't already hidden
        if '#0' in self.tree['columns'] and self.tree.column('#0', 'width') != 0 :
             self.tree.column('#0', width=0, stretch=tk.NO)

    def color_rows(self):
        """
        Applies background and foreground colors to Treeview rows based on their 'Status'.
        Uses color settings from config.STATUS_COLORS.
        """
        if not self.tree or self.app.status_df is None or self.app.status_df.empty: return

        # Define styles (tags) for different statuses using colors from config
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

        # Status categories for coloring
        action_statuses = ["Ready to order", "Permit", "Waiting Measure"] 
        good_statuses = ["Ready to dispatch", "In install", "Done", "Waiting for materials"]
        status_col_name = "Status" 
        
        try: 
            # Get the index of the 'Status' column in the Treeview
            status_column_tree_index = config.EXPECTED_COLUMNS.index(status_col_name)
        except ValueError:
            logging.error(f"DMT: '{status_col_name}' column not found in EXPECTED_COLUMNS. Cannot color rows.")
            return

        # Iterate over each row in the Treeview and apply the appropriate style tag
        for item_id in self.tree.get_children():
            try:
                # Preserve the original DataFrame index tag
                df_index_tag_list = self.tree.item(item_id, "tags")
                df_index_tag = df_index_tag_list[0] if df_index_tag_list else f"err_idx_{item_id}" # Fallback tag
                new_tags_for_item = [df_index_tag] # Start with the df_index tag

                # Get the status value from the Treeview row
                values = self.tree.item(item_id, "values")
                if values and len(values) > status_column_tree_index:
                    status = str(values[status_column_tree_index])
                    # Append the corresponding style tag based on status
                    if status == "New": new_tags_for_item.append("new_style")
                    elif status == "Closed" or status == "Cancelled/Postponed": new_tags_for_item.append("closed_style")
                    elif status == config.REVIEW_MISSING_STATUS: new_tags_for_item.append("review_missing_style")
                    elif status in action_statuses: new_tags_for_item.append("action_needed_style")
                    elif status in good_statuses: new_tags_for_item.append("all_good_style")
                    else: new_tags_for_item.append("default_status_style")
                else:
                    # If status can't be determined, apply default style
                    new_tags_for_item.append("default_status_style")
                
                # Set the combined tags for the item (df_index + style_tag)
                self.tree.item(item_id, tags=tuple(new_tags_for_item))
            except Exception as e:
                # Log error and apply a fallback style to prevent crashes
                logging.error(f"DMT: Error coloring row {item_id}: {e}", exc_info=True)
                df_idx_tag_fallback = self.tree.item(item_id,"tags")[0] if self.tree.item(item_id,"tags") else f"err_idx_fb_{item_id}"
                self.tree.item(item_id, tags=(df_idx_tag_fallback, "default_status_style"))

    def sort_treeview_column(self, col, reverse):
        """
        Sorts the Treeview rows based on the values in the specified column.
        Handles numeric, date, and string sorting.
        Args:
            col: The name of the column to sort by.
            reverse: Boolean, True for descending sort, False for ascending.
        """
        if not self.tree or self.app.status_df is None or self.app.status_df.empty: return

        date_columns = ['Order Date', 'Turn in Date'] # Columns to be treated as dates for sorting

        try:
            # Prepare a list of (sort_value, item_id) tuples
            items_to_sort = []
            for item_id in self.tree.get_children(''):
                sort_value = None
                tags = self.tree.item(item_id, "tags") # Get DataFrame index
                
                if not tags: 
                    # Fallback: if no df_index tag, sort by the displayed value (less accurate for some types)
                    logging.warning(f"DMT: Cannot sort item {item_id}, missing df_index tag. Using displayed value.")
                    sort_value = self.tree.set(item_id, col) 
                else:
                    try:
                        df_index = int(tags[0])
                        if df_index not in self.app.status_df.index: 
                             # Fallback if df_index is invalid
                             logging.warning(f"DMT: Invalid df_index {df_index} for item {item_id} during sort. Using displayed value.")
                             sort_value = self.tree.set(item_id, col)
                        else:
                            # Get the original value from the DataFrame for accurate sorting
                            original_value = self.app.status_df.loc[df_index, col]
                            
                            # Convert to appropriate type for sorting
                            if col in date_columns:
                                sort_value = pd.to_datetime(original_value, errors='coerce') # NaT for unparseable dates
                            elif col in config.CURRENCY_COLUMNS:
                                try: 
                                    sort_value = float(str(original_value).replace('$', '').replace(',', ''))
                                except (ValueError, TypeError): sort_value = str(original_value) # Fallback to string sort
                            else: # For other columns, try numeric then string
                                try: sort_value = float(original_value)
                                except (ValueError, TypeError): sort_value = str(original_value)
                    except (IndexError, KeyError, ValueError, TypeError) as e:
                        # General fallback for any error during value retrieval/conversion
                        logging.warning(f"DMT: Sort fallback for col {col}, item {item_id} (df_idx {tags[0]}): {e}. Using displayed value.")
                        sort_value = self.tree.set(item_id, col) 
                
                items_to_sort.append((sort_value, item_id))

            # Custom sort key to handle different data types and NaN/NaT values
            def sort_key(item_tuple):
                value = item_tuple[0]
                if isinstance(value, pd.Timestamp): return (0, value) # Dates first
                if pd.isna(value): return (1, pd.Timestamp.min if not reverse else pd.Timestamp.max) # Handle NaNs/NaTs
                if isinstance(value, (int, float)): return (2, value) # Numbers next
                return (3, str(value).lower()) # Strings last (case-insensitive)

            items_to_sort.sort(key=sort_key, reverse=reverse) # Perform the sort
            
            # Reorder items in the Treeview
            for index, (val, item_id) in enumerate(items_to_sort):
                self.tree.move(item_id, '', index)
                
            # Update the column heading to toggle sort direction on next click
            self.tree.heading(col, command=lambda _col=col: self.sort_treeview_column(_col, not reverse))
        except Exception as e:
            logging.error(f"DMT: Error sorting column {col}: {e}", exc_info=True)

    def on_tab_selected(self):
        """
        Called when this tab is selected in the notebook.
        Can be used to refresh data or UI elements if needed.
        """
        logging.info("Data Management tab selected.")
        if self.tree: 
            # If tree is empty but data exists, populate it (e.g., first time tab is shown after data load)
            if not self.tree.get_children() and self.app.status_df is not None and not self.app.status_df.empty:
                logging.info("Data Management tab was empty, populating treeview on selection.")
                self.populate_treeview()
            else: 
                # Otherwise, just ensure columns and widths are up-to-date
                self.configure_treeview_columns()
                self.set_column_widths_from_preferred()
        else:
            logging.warning("Data Management tab selected, but treeview not initialized.")
