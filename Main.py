import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# --- Configuration Variables ---
STATUS_FILE = "invoice_status.pkl"
OUTPUT_FILE = "open_invoices.xlsx"
EXPECTED_COLUMNS = [
    '#', 'Invoice #', 'Order Date', 'Turn in Date', 'Account',
    'Invoice Total', 'Balance', 'Salesperson', 'Project Coordinator',
    'Status', 'Notes'
]
DEFAULT_FONT = ('Calibri', 14)
DEFAULT_PADDING = 10
DATE_FORMAT = '%b-%d'
CURRENCY_COLUMNS = ['Invoice Total', 'Balance']
CURRENCY_FORMAT = '${:,.2f}'
# --- End Configuration ---

def load_excel(excel_file):
    """Loads an Excel file into a Pandas DataFrame."""
    try:
        df = pd.read_excel(excel_file)
        return df
    except FileNotFoundError:
        messagebox.showerror("Error", f"File not found: {excel_file}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel: {e}")
        return None

def load_status():
    """Loads the current status DataFrame from a pickle file."""
    try:
        df = pd.read_pickle(STATUS_FILE)
        return df
    except FileNotFoundError:
        # Create empty DataFrame if file doesn't exist
        df = pd.DataFrame(columns=EXPECTED_COLUMNS)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error loading status: {e}")
        return None

def save_status(df):
    """Saves the current status DataFrame to a pickle file."""
    try:
        df.to_pickle(STATUS_FILE)
        messagebox.showinfo("Info", "Status saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving status: {e}")

def process_data(new_df, status_df):
    """
    Processes the new Excel data against the current status.
    """
    # 1. Sanitize column names in new_df
    new_df.columns = [str(col).strip() for col in new_df.columns]
    new_df.columns = [col.replace('\n', '').replace('\r', '') for col in new_df.columns]

    # 2. Merge DataFrames
    merged_df = new_df.merge(status_df, on='Invoice #', how='outer', indicator=True, suffixes=('_new', '_old'))

    def update_row(row):
        if row['_merge'] == 'both':
            for col in new_df.columns:
                if col != 'Invoice #':
                    new_col_name = col + '_new'
                    old_col_name = col + '_old'
                    if new_col_name in row:
                        row[col] = row[new_col_name]
                    elif old_col_name in row:
                        row[col] = row[old_col_name]
        elif row['_merge'] == 'left_only':
            for col in new_df.columns:
                if col != 'Invoice #':
                    new_col_name = col + '_new'
                    if new_col_name in row:
                        row[col] = row[new_col_name]
                    else:
                        row[col] = None
            row['Status'] = 'New'
            row['Notes'] = ''
        elif row['_merge'] == 'right_only':
            row = row.copy()
            row['Status'] = 'Closed'
        return row

    merged_df = merged_df.apply(update_row, axis=1)

    # 3. Clean up the merged DataFrame
    cols_to_drop = ['_merge']
    for col in new_df.columns:
        if col != 'Invoice #':
            cols_to_drop.append(col + '_new')
    for col in status_df.columns:
        if col not in ('Invoice #', 'Status', 'Notes'):
            cols_to_drop.append(col + '_old')

    final_df = merged_df.drop(columns=cols_to_drop, errors='ignore')

    # 4. Ensure final DataFrame has the expected columns
    for col in EXPECTED_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = None

    return final_df[EXPECTED_COLUMNS]

def generate_report(df):
    """Generates a report of open invoices."""
    open_invoices = df[df['Status'] != 'Closed']
    try:
        open_invoices.to_excel(OUTPUT_FILE, index=False)
        messagebox.showinfo("Info", f"Report generated: {OUTPUT_FILE}")
    except Exception as e:
        messagebox.showerror("Error", f"Error generating report: {e}")

class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Invoice Status Tracker")
        self.maximize_window()

        self.style = ttk.Style(self)
        self.style.configure("Treeview", font=DEFAULT_FONT)
        self.style.configure("Treeview.Heading", font=DEFAULT_FONT)

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
                    value = CURRENCY_FORMAT.format(value)
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
                self.color_rows()  # Re-apply row colors
                self.populate_treeview()
                # self.color_rows()

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
        generate_report(self.status_df)

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

if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()