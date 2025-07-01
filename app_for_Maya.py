import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pandastable import Table
from pandasql import sqldf
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

class ExcelTableApp:
    """
    GUI Application for loading, aggregating, and saving Excel data using SQL logic.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("App for my dear auntie Maya")
        self.root.geometry("800x600")

        # Initialize core components
        self.df = pd.DataFrame()
        self.table = None
        self.current_file = None

        self._check_dependencies()
        self._setup_widgets()
        self._initialize_table()

        logging.info("Application initialized successfully.")

    def _check_dependencies(self):
        """Ensure required libraries are available."""
        try:
            import openpyxl
            import xlwt
            import pandasql
            import pandastable
        except ImportError as e:
            messagebox.showerror(
                "Missing Dependency",
                f"Missing required library: {e}. Please install pandas, openpyxl, pandastable, pandasql, and xlwt."
            )
            raise

    def _setup_widgets(self):
        """Configure GUI layout and buttons."""
        self.file_label = tk.Label(self.root, text="No file loaded", pady=10)
        self.file_label.pack()

        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=5)

        tk.Button(self.button_frame, text="Load Excel File", command=self.load_file).pack(side=tk.LEFT, padx=5)

        self.refresh_button = tk.Button(self.button_frame, text="Refresh", command=self.refresh_file, state=tk.DISABLED)
        self.refresh_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(self.button_frame, text="Save to Excel", command=self.save_file, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)

    def _setup_filters(self):
        """Create UI filter widgets above the table."""
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=5)

        tk.Label(self.filter_frame, text="Filter by Name:").pack(side=tk.LEFT)
        self.name_filter_entry = tk.Entry(self.filter_frame, width=15)
        self.name_filter_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(self.filter_frame, text="Filter by Subject:").pack(side=tk.LEFT)
        self.subject_filter_entry = tk.Entry(self.filter_frame, width=15)
        self.subject_filter_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(self.filter_frame, text="Apply Filter", command=self.apply_filter).pack(side=tk.LEFT, padx=10)

    def _initialize_table(self):
        """Display empty table on startup."""
        self.table = Table(self.table_frame, dataframe=self.df, showtoolbar=False, showstatusbar=True)
        self.table.show()

    def apply_aggregation(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Apply predefined SQL aggregation query to filter data.
        """
        query = (
            """
            SELECT * FROM df df1 
            WHERE NOT EXISTS (
            SELECT 1 FROM df df2 
            WHERE df1.name = df2.name AND df2.Rev IN ('A01', 'V01', 'X01')) 
            AND df1.Subject IN ('EXECUTE', 'DEFINE/EXECUTE')
            """
        )
        try:
            result_df = sqldf(query, {"df": df})
            return result_df if not result_df.empty else df
        except Exception as e:
            logging.warning(f"Aggregation failed: {e}")
            messagebox.showwarning("Warning", f"Aggregation failed: {e}. Showing original data.")
            return df

    def apply_filter(self):
        """Apply user-entered filters to the DataFrame."""
        if self.df is None or self.df.empty:
            messagebox.showinfo("Info", "No data to filter.")
            return

        filtered_df = self.df.copy()

        name_filter = self.name_filter_entry.get().strip().lower()
        subject_filter = self.subject_filter_entry.get().strip().lower()

        if name_filter:
            filtered_df = filtered_df[filtered_df['name'].astype(str).str.lower().str.contains(name_filter)]

        if subject_filter:
            filtered_df = filtered_df[filtered_df['Subject'].astype(str).str.lower().str.contains(subject_filter)]

        if filtered_df.empty:
            messagebox.showinfo("Info", "No matching records found.")
        else:
            self.table.model.df = filtered_df
            self.table.redraw()

    def load_file(self):
        """Load Excel file and apply aggregation."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
            self.current_file = file_path
            self.df = self.apply_aggregation(self.df)

            self._update_table()
            self._update_status(f"Aggregated data from: {os.path.basename(file_path)}")

            self.refresh_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

        except Exception as e:
            logging.error(f"File loading error: {e}")
            messagebox.showerror("Error", f"Could not load file: {e}")
            self.df = pd.DataFrame()

    def refresh_file(self):
        """Re-load the last loaded file and re-apply aggregation."""
        if not self.current_file:
            return

        try:
            self.df = pd.read_excel(self.current_file)
            self.df = self.apply_aggregation(self.df)
            self._update_table()
            self._update_status(f"Aggregated data from: {os.path.basename(self.current_file)}")
        except Exception as e:
            logging.error(f"Refresh error: {e}")
            messagebox.showerror("Error", f"Could not refresh file: {e}")

    def save_file(self):
        """Save the aggregated DataFrame to Excel."""
        if self.df.empty:
            messagebox.showerror("Error", "No data to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Save Aggregated Data"
        )

        if not file_path:
            return

        try:
            if file_path.endswith(".xls"):
                self.df.to_excel(file_path, index=False, engine="xlwt")
            else:
                self.df.to_excel(file_path, index=False)

            self._update_status(f"Data saved to: {os.path.basename(file_path)}")
        except Exception as e:
            logging.error(f"Save error: {e}")
            messagebox.showerror("Error", f"Could not save file: {e}")

    def _update_table(self):
        """Update table view with new DataFrame."""
        self.table.model.df = self.df
        self.table.redraw()

    def _update_status(self, message: str):
        """Update file label with the given message."""
        self.file_label.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTableApp(root)
    root.mainloop()
