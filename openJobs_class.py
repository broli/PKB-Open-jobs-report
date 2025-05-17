import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar, OptionMenu
import sv_ttk # Import the sun valley ttk theme
import pandas as pd
from data_utils import load_status, save_status, load_excel, process_data
import logging

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
OUTPUT_FILE = "open_invoices.xlsx"
EXPECTED_COLUMNS = [
    '#', 'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]
DEFAULT_FONT = ('Calibri', 16)
DEFAULT_PADDING = 10
DATE_FORMAT = '%b-%d'
CURRENCY_COLUMNS = ['Invoice Total', 'Balance']
CURRENCY_FORMAT = '${:,.2f}'
ALLOWED_STATUS = [
    "Waiting Measure", "Ready to order", "Waiting for materials",
    "Ready to dispatch", "In install", "Done", "Permit",
    "Cancelled/Postponed"
]
# Configure the logging level
logging.basicConfig(level=logging.DEBUG)  # Set to logging.INFO or logging.ERROR for production
#logging.basicConfig(level=logging.ERROR)  # Set to logging.INFO or logging.ERROR for production
#

class OpenJobsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        sv_ttk.set_theme("dark")  # Or "light"
        self.title("Open Jobs App")  # Application Title
        self.maximize_window()

        self.style = ttk.Style(self)
        self.style.configure("Treeview", font=DEFAULT_FONT)
        self.style.configure("Treeview.Heading", font=DEFAULT_FONT)

        #app background color
        self.configure(background="steel blue")

        self.status_df = load_status()
        self.create_widgets()

    def maximize_window(self):
        """Maximizes the application window."""
        self.state('zoomed')

    def create_widgets(self):
        # --- Menu Bar ---
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load New Excel", command=self.load_new_excel)
        file_menu.add_command(label="Save", command=self.save_data)
        file_menu.add_command(label="Generate Report", command=self.generate_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menu_bar)

        # --- Treeview for Displaying Data ---
        self.tree = ttk.Treeview(self, columns=list(self.status_df.columns), show="headings")
        for col in self.status_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)  # Initial width, left-aligned
        self.tree.pack(fill=tk.BOTH, expand=True)

        # --- Apply Gridlines Style ---
        self.apply_gridline_style()  # Apply gridlines initially

        # --- Editing Functionality ---
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected_row)
        self.editing_window = None

        self.populate_treeview()
        self.set_column_widths()
        self.color_rows()
        self.apply_gridline_style()  # Apply gridlines again just in case


    def apply_gridline_style(self):
        """Applies the gridline style to the Treeview."""
        self.style.layout("Treeview", [('Treeview.field', {'sticky': 'nswe'})])

    def configure_treeview_columns(self):
        """
        Completely rebuilds the Treeview columns to match the current data.
        This is the most robust way to ensure the Treeview's structure is correct.
        """

        # 1. Get all current items and their values
        items = self.tree.get_children()
        all_data = [self.tree.item(item, 'values') for item in items]

        # 2. Detach the Treeview from its parent
        self.tree.pack_forget()

        # 3. Create a new Treeview
        self.tree = ttk.Treeview(self, columns=list(self.status_df.columns), show="headings")

        # 4. Re-pack the Treeview
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 5. Re-apply the style and bindings
        self.apply_gridline_style()  # Re-apply gridline style
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected_row)

        # 6. Recreate the headings
        for col in self.status_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)  # Initial width, left-aligned

        # 7. Re-insert the data (if any)
        for values in all_data:
            self.tree.insert("", tk.END, values=values)

    def populate_treeview(self):
        """Populates the Treeview with data."""
        date_columns = ['Order Date', 'Turn in Date']
        for i in self.tree.get_children():
            self.tree.delete(i)

        for index, row in self.status_df.iterrows():
            values = []
            for col in self.status_df.columns:
                value = row[col]
                if col in date_columns and pd.notna(value):
                    value = value.strftime(DATE_FORMAT)
                if col in CURRENCY_COLUMNS and pd.notna(value):
                    if pd.notna(value):
                        try:
                            value = float(value)  # Convert to float if possible
                            value = CURRENCY_FORMAT.format(value)
                        except ValueError:
                            # If conversion fails, keep the original value
                            pass
                values.append(value)
            self.tree.insert("", tk.END, values=values)

    def load_new_excel(self):
        """Loads a new Excel file and processes the data."""

        excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if excel_file_path:
            new_df = load_excel(excel_file_path)
            if new_df is not None:
                self.status_df = process_data(new_df, self.status_df)
                self.configure_treeview_columns()
                self.after(10, self.set_column_widths)
                self.color_rows()
                self.populate_treeview()

    def save_data(self):
        """Saves the current data."""

        all_data = []
        for item in self.tree.get_children():
            all_data.append(self.tree.item(item, 'values'))
        self.status_df = pd.DataFrame(all_data, columns=list(self.status_df.columns))
        save_status(self.status_df)

    def generate_report(self):
        """Generates the report."""

        self.save_data()
        self.generate_report_helper()  # Call the helper function

    def generate_report_helper(self):
        """Helper function to generate the report."""
        open_invoices = self.status_df[self.status_df['Status'] != 'Closed']
        try:
            open_invoices.to_excel(OUTPUT_FILE, index=False)
            messagebox.showinfo("Info", f"Report generated: {OUTPUT_FILE}")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating report: {e}")

    
    def on_double_click(self, event):
        """Handles double-clicks on Treeview items to allow editing."""
        item = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)  # e.g., "#1", "#2"

        if item and column_id_str:  # Check if a valid row and column were clicked
            try:
                # Convert the column ID string (e.g., "#3") to a 0-based index
                column_index = int(column_id_str.replace("#", "")) - 1

                # Get the actual name of the column from your DataFrame columns list
                if 0 <= column_index < len(self.status_df.columns):
                    actual_column_name = self.status_df.columns[column_index]

                    if self.editing_window: # Destroy any existing editing window
                        self.editing_window.destroy()
                        self.editing_window = None

                    #  Direct to specific editor based on column name
                    if actual_column_name == "Status":
                        logging.debug(f"Double-clicked 'Status' column. Item: {item}, Column ID: {column_id_str}")
                        self.create_status_dropdown(item, column_id_str) # Use column_id_str for this function
                    elif actual_column_name == "Notes":
                        logging.debug(f"Double-clicked 'Notes' column. Item: {item}, Column Name: {actual_column_name}")
                        self.create_editing_window(item, actual_column_name) # Use actual_column_name
                    else:
                        logging.debug(f"Double-clicked column '{actual_column_name}' is not configured for editing or has no special editor.")
                else:
                    logging.warning(f"Invalid column index derived: {column_index}")

            except ValueError:
                logging.error(f"Could not parse column ID: {column_id_str}")
            except IndexError:
                logging.error(f"Column index out of bounds: {column_index}") # Should use the derived column_index in the log
        else:
            logging.debug("Double-click was not on a valid cell.")


    
    def create_editing_window(self, item, actual_column_name): # Parameter renamed for clarity
        """Creates a pop-up window for editing a cell."""
        logging.debug(f"Starting create_editing_window for item: {item}, column: {actual_column_name}")
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {actual_column_name}") # Use the actual column name

        # Get the index of the column using the actual_column_name
        try:
            column_idx = list(self.status_df.columns).index(actual_column_name)
        except ValueError:
            logging.error(f"Column name '{actual_column_name}' not found in DataFrame columns.")
            self.editing_window.destroy() # Close the Toplevel window if column is invalid
            return

        current_value = self.tree.item(item, "values")[column_idx]

        entry = tk.Entry(self.editing_window, width=50)
        entry.insert(tk.END, str(current_value)) # Ensure current_value is a string for the entry
        entry.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        def save_edit():
            new_value = entry.get()
            # current_tree_values is a tuple, convert to list for modification
            current_tree_values_list = list(self.tree.item(item, "values"))

            # Use the same column_idx derived earlier
            current_tree_values_list[column_idx] = new_value

            # Update the tree item with the modified list (converted back to tuple)
            self.tree.item(item, values=tuple(current_tree_values_list))

            self.editing_window.destroy()
            self.editing_window = None
            self.color_rows() # Assuming this function relies on updated tree values
            # self.set_column_widths() # Consider if this is needed immediately after a single cell edit

        save_button = ttk.Button(self.editing_window, text="Save", command=save_edit)
        save_button.pack(pady=DEFAULT_PADDING)


    def create_status_dropdown(self, item, column): # 'column' here is column_id_str e.g., "#10"
        """Creates a dropdown menu for editing the Status."""
        self.editing_window = tk.Toplevel(self)
        
        column_index = -1 # Initialize to an invalid index
        try:
            # Correctly get the column name and index from the Treeview column identifier
            column_index = int(column[1:]) - 1  # convert to int and -1 for 0 index
            column_name = self.status_df.columns[column_index] # This will be "Status"
        except (ValueError, IndexError) as e:
            logging.error(f"Error deriving column index/name from ID '{column}': {e}")
            self.editing_window.destroy()
            return

        self.editing_window.title(f"Edit {column_name}") 

        current_value_tuple = self.tree.item(item, "values")
        
        if not (0 <= column_index < len(current_value_tuple)):
            logging.error(f"Column index {column_index} is out of bounds for item values (length {len(current_value_tuple)}).")
            self.editing_window.destroy()
            return
            
        current_value = current_value_tuple[column_index]

        status_var = StringVar(self.editing_window)
        
        if str(current_value) in ALLOWED_STATUS: # Ensure comparison with string form if necessary
            status_var.set(current_value)
        elif ALLOWED_STATUS: 
            status_var.set(ALLOWED_STATUS[0])
        # else: status_var will be empty, which is fine for OptionMenu, or set a default

        status_dropdown = OptionMenu(self.editing_window, status_var, *ALLOWED_STATUS)
        status_dropdown.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        def save_edit():
            new_value = status_var.get()
            
            # 1. Get the current values from the tree item (it's a tuple).
            current_tree_values_tuple = self.tree.item(item, "values")
            # 2. Convert this tuple to a list to allow modification.
            values_list = list(current_tree_values_tuple)
            
            # 3. Modify the list at the correct column_index.
            # Ensure column_index is still valid (it should be from the outer scope)
            if 0 <= column_index < len(values_list):
                values_list[column_index] = new_value
            else:
                logging.error(f"column_index {column_index} out of bounds during save_edit.")
                self.editing_window.destroy()
                self.editing_window = None
                return

            # 4. Update the tree item with the modified list (converted back to a tuple).
            self.tree.item(item, values=tuple(values_list))
            
            self.editing_window.destroy()
            self.editing_window = None 
            self.color_rows()
            self.set_column_widths() # Optional: if status change affects width significantly

        save_button = ttk.Button(self.editing_window, text="Save", command=save_edit)
        save_button.pack(pady=DEFAULT_PADDING)



    def delete_selected_row(self, event):
        """Deletes the currently selected row from the Treeview."""

        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item)
            self.color_rows()
            self.set_column_widths()

    def set_column_widths(self):
        """Adjusts column widths to fit content."""

        for col in self.status_df.columns:
            self.tree.column(col, width=0, stretch=tk.NO)
            for item in self.tree.get_children():
                try:  # Add try-except block
                    value = self.tree.item(item, "values")[list(self.status_df.columns).index(col)]
                    col_width = max(self.tree.column(col, 'width'), len(str(value)) * 7 + 30)
                    self.tree.column(col, width=col_width, anchor=tk.W)  # Left-align
                except ValueError:
                    # Handle the case where the column might not exist in this item
                    pass
            self.tree.column('#0', width=0)

    def color_rows(self):
        """Colors rows based on status."""

        for item in self.tree.get_children():
            try:
                status = self.tree.item(item, "values")[list(self.status_df.columns).index("Status")]
                if status == "Closed":
                    self.tree.tag_configure("closed", background="lightgray")
                    self.tree.item(item, tags="closed")
                elif status == "New":
                    self.tree.tag_configure("new", background="lightblue")
                    self.tree.item(item, tags="new")
                else:
                    self.tree.tag_configure("open", background="white")
                    self.tree.item(item, tags="open")
            except IndexError:
                pass