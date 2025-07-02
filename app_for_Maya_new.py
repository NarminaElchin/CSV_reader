import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pandastable import Table
from pandasql import sqldf
import os
import logging
from tkcalendar import DateEntry
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

class ExcelTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer & Aggregator")
        self.root.state('zoomed')  # Fullscreen

        self.df = pd.DataFrame()
        self.table = None
        self.current_file = None
        self.filters = {}

        self._setup_widgets()
        self._setup_filters()
        self._initialize_table()

    def _setup_widgets(self):
        self.file_label = tk.Label(self.root, text="No file loaded", pady=10, font=("Arial", 12))
        self.file_label.pack()

        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=5)

        tk.Button(self.button_frame, text="Load Excel File", command=self.load_file).pack(side=tk.LEFT, padx=5)
        self.refresh_button = tk.Button(self.button_frame, text="Refresh", command=self.refresh_file, state=tk.DISABLED)
        self.refresh_button.pack(side=tk.LEFT, padx=5)
        self.save_button = tk.Button(self.button_frame, text="Save to Excel", command=self.save_file, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

    def _on_date_validate(self, action, value_if_allowed):
        # Allow user to delete date field manually
        return True

    def _setup_filters(self):
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=10)

        def add_dropdown(label, col_name):
            tk.Label(self.filter_frame, text=label).pack(side=tk.LEFT, padx=5)
            combo = ttk.Combobox(self.filter_frame, state="readonly", width=20)
            combo.pack(side=tk.LEFT)
            self.filters[col_name] = combo

        add_dropdown("Subject:", "Subject")
        add_dropdown("Rev:", "Rev")
        add_dropdown("Revision Desc:", "Revision Description")
        add_dropdown("UBOC Code:", "UBOC Return Code")

        # DateEntry with blank default
        vcmd = (self.root.register(self._on_date_validate), '%d', '%P')

        tk.Label(self.filter_frame, text="UBOC Approval Date From:").pack(side=tk.LEFT, padx=5)
        self.date_from = DateEntry(self.filter_frame, date_pattern='dd/mm/yyyy', width=12,
                                   validate='key', validatecommand=vcmd)
        self.date_from.pack(side=tk.LEFT)
        self.date_from.delete(0, tk.END)

        tk.Label(self.filter_frame, text="To:").pack(side=tk.LEFT)
        self.date_to = DateEntry(self.filter_frame, date_pattern='dd/mm/yyyy', width=12,
                                 validate='key', validatecommand=vcmd)
        self.date_to.pack(side=tk.LEFT)
        self.date_to.delete(0, tk.END)

        tk.Button(self.filter_frame, text="Apply Filters", command=self.apply_filter).pack(side=tk.LEFT, padx=10)

    def _initialize_table(self):
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)
        self.table = Table(self.table_frame, dataframe=self.df, showtoolbar=False, showstatusbar=True)
        self.table.show()

    def apply_aggregation(self, df: pd.DataFrame) -> pd.DataFrame:
        query = (
            "SELECT * FROM df df1 "
            "WHERE NOT EXISTS ("
            "SELECT 1 FROM df df2 "
            "WHERE df1.name = df2.name AND df2.Rev IN ('V01', 'X01')) "
            "AND df1.Rev != 'A01'"
            "AND df1.Subject IN ('EXECUTE', 'DEFINE/EXECUTE')"
        )
        try:
            result_df = sqldf(query, {"df": df})
            return result_df if not result_df.empty else df
        except Exception as e:
            messagebox.showwarning("Warning", f"Aggregation failed: {e}. Showing original data.")
            return df

    def apply_filter(self):
        if self.df.empty:
            messagebox.showinfo("Info", "No data to filter.")
            return

        filtered_df = self.df.copy()

        # Dropdown filters
        for col, combo in self.filters.items():
            value = combo.get()
            if value == "BLANK":
                # Filter rows where column is NaN or empty string (after stripping spaces)
                filtered_df = filtered_df[filtered_df[col].isna() | (filtered_df[col].str.strip() == "")]
            elif value and value != "ALL":
                filtered_df = filtered_df[filtered_df[col] == value]

        # Date filtering
        date_col = "UBOC Approval Date"
        if date_col in filtered_df.columns:
            try:
                from_date = self.date_from.get()
                to_date = self.date_to.get()

                if from_date.strip() or to_date.strip():
                    from_dt = self.date_from.get_date() if from_date.strip() else None
                    to_dt = self.date_to.get_date() if to_date.strip() else None

                    date_series = pd.to_datetime(filtered_df[date_col], errors='coerce').dt.date
                    valid_dates = date_series.notnull()

                    mask = pd.Series(True, index=filtered_df.index)
                    if from_dt:
                        mask &= date_series >= from_dt
                    if to_dt:
                        mask &= date_series <= to_dt

                    filtered_df = filtered_df[valid_dates & mask]
                # else: no filtering, keep all

            except Exception as e:
                messagebox.showerror("Error", f"Date filtering failed: {e}")
                return

        if filtered_df.empty:
            messagebox.showinfo("Info", "No matching records found.")
        else:
            self.table.model.df = filtered_df
            self.table.redraw()

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path, dtype=str)
            self.current_file = file_path
            self.df = self.apply_aggregation(self.df)

            self._update_table()
            self._update_status(f"Aggregated data from: {os.path.basename(file_path)}")

            self.refresh_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

            self._populate_filter_dropdowns()

        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {e}")
            self.df = pd.DataFrame()

    def _populate_filter_dropdowns(self):
        for col, combo in self.filters.items():
            if col in self.df.columns:
                unique_values = sorted(self.df[col].dropna().unique())
                combo['values'] = ["ALL"] + unique_values + ["BLANK"]
                combo.set("ALL")

        self.date_from.delete(0, tk.END)
        self.date_to.delete(0, tk.END)

    def refresh_file(self):
        if not self.current_file:
            return

        try:
            self.df = pd.read_excel(self.current_file, dtype=str)
            self.df = self.apply_aggregation(self.df)
            self._update_table()
            self._update_status(f"Aggregated data from: {os.path.basename(self.current_file)}")
            self._populate_filter_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Could not refresh file: {e}")

    def save_file(self):
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
            messagebox.showerror("Error", f"Could not save file: {e}")

    def _update_table(self):
        self.table.model.df = self.df
        self.table.redraw()

    def _update_status(self, message: str):
        self.file_label.config(text=message)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTableApp(root)
    root.mainloop()
