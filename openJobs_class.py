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
        column = self.tree.identify_column(event.x)

        if item and column in ("#Status", "#Notes"):
            if self.editing_window:
                self.editing_window.destroy()
            self.create_editing_window(item, column)

    def create_editing_window(self, item, column):
        """Creates a pop-up window for editing a cell."""
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {column[1:]}")

        current_value = self.tree.item(item, "values")[list(self.status_df.columns).index(column[1:])]
        entry = tk.Entry(self.editing_window, width=50)
        entry.insert(tk.END, current_value)
        entry.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        def save_edit():
            new_value = entry.get()
            values = self.tree.item(item, "values")
            values[list(self.status_df.columns).index(column[1:])] = new_value
            self.tree.item(item, values=values)
            self.editing_window.destroy()
            self.editing_window = None
            self.color_rows()
            self.set_column_widths()

        save_button = ttk.Button(self.editing_window, text="Save", command=save_edit)
        save_button.pack(pady=DEFAULT_PADDING)

    def create_status_dropdown(self, item, column):
        """Creates a dropdown menu for editing the Status."""
        self.editing_window = tk.Toplevel(self)
        self.editing_window.title(f"Edit {column[1:]}")  # Remove the '#' here

        # Correctly get the column name from the Treeview column identifier
        column_index = int(column[1:]) - 1  # convert to int and -1 for 0 index
        column_name = self.status_df.columns[column_index]
        current_value = self.tree.item(item, "values")[column_index]

        status_var = StringVar(self.editing_window)
        status_var.set(current_value)  # Set the current value

        status_dropdown = OptionMenu(self.editing_window, status_var, *ALLOWED_STATUS)
        status_dropdown.pack(padx=DEFAULT_PADDING, pady=DEFAULT_PADDING)

        def save_edit():
            new_value = status_var.get()
            if new_value in ALLOWED_STATUS:
                values = self.tree.item(item, "values")
                values[column_index] = new_value
                self.tree.item(item, values=values)
                self.editing_window.destroy()
                self.editing_window = None
                self.color_rows()
                self.set_column_widths()
            else:
                messagebox.showerror("Error", "Invalid Status")

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