import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pandastable import Table
from pandasql import sqldf
import os
from tkcalendar import DateEntry
from datetime import datetime

class ExcelTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer & Aggregator")
        self.root.state('zoomed')

        self.df_raw_original = pd.DataFrame()
        self.df_raw = pd.DataFrame()
        self.current_file = None
        self.selected_uboc_values = []

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
        return True

    def _setup_filters(self):
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=10)

        uboc_frame = tk.Frame(self.filter_frame)
        uboc_frame.pack(side=tk.LEFT, padx=5)
        tk.Label(uboc_frame, text="UBOC Code:").pack()
        self.uboc_button_var = tk.StringVar(value="Select UBOC Code ▼")
        self.uboc_dropdown_button = tk.Button(uboc_frame, textvariable=self.uboc_button_var, width=25,
                                              command=self.show_uboc_dropdown)
        self.uboc_dropdown_button.pack()

        vcmd = (self.root.register(self._on_date_validate), '%d', '%P')
        tk.Label(self.filter_frame, text="Workflow Start From:").pack(side=tk.LEFT, padx=5)
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

    def show_uboc_dropdown(self):
        if "UBOC Return Code" not in self.df_raw.columns:
            return

        top = tk.Toplevel(self.root)
        top.title("Select UBOC Code")
        top.geometry("+%d+%d" % (self.root.winfo_pointerx(), self.root.winfo_pointery()))
        top.grab_set()

        self.uboc_vars = {}
        all_values = sorted(self.df_raw["UBOC Return Code"].dropna().unique())
        all_values = ["ALL"] + all_values + ["BLANK"]

        for val in all_values:
            var = tk.BooleanVar(value=val in self.selected_uboc_values)
            cb = tk.Checkbutton(top, text=val, variable=var, anchor="w")
            cb.pack(fill="x", padx=10)
            self.uboc_vars[val] = var

        def confirm_selection():
            selected = [k for k, v in self.uboc_vars.items() if v.get()]
            self.selected_uboc_values = selected
            display_text = ", ".join(selected[:3])
            if len(selected) > 3:
                display_text += f"... (+{len(selected)-3})"
            self.uboc_button_var.set(display_text or "Select UBOC Code ▼")
            top.destroy()

        tk.Button(top, text="Confirm", command=confirm_selection).pack(pady=5)

    def _initialize_table(self):
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)
        self.table = Table(self.table_frame, dataframe=self.df_raw, showtoolbar=False, showstatusbar=True)
        self.table.show()

    def apply_initial_base_sql(self, df):
        query = (
            "SELECT * FROM df df1 "
            "WHERE NOT EXISTS ("
            "SELECT 1 FROM df df2 WHERE df1.name = df2.name AND df2.Rev IN ('V01', 'X01')) "
            "AND df1.Rev != 'A01' "
            "AND df1.Subject IN ('EXECUTE', 'DEFINE/EXECUTE') "
        )
        try:
            return sqldf(query, {"df": df})
        except Exception as e:
            messagebox.showerror("SQL Error", f"Base SQL filtering failed: {e}")
            return df

    def apply_aggregation(self, df):
        query = (
            "SELECT "
            "df1.Name, df1.Title, "
            "SUM(CASE WHEN df1.Rev LIKE 'B%' THEN 1 ELSE 0 END) AS IFR, "
            "SUM(CASE WHEN df1.Rev LIKE 'D%' THEN 1 ELSE 0 END) AS AFD, "
            "SUM(CASE WHEN df1.Rev LIKE 'U%' THEN 1 ELSE 0 END) AS AFU, "
            "SUM(CASE WHEN df1.Rev LIKE 'H%' THEN 1 ELSE 0 END) AS AFH, "
            "SUM(CASE WHEN df1.Rev LIKE 'I%' THEN 1 ELSE 0 END) AS IFI, "
            "SUM(CASE WHEN df1.Rev LIKE 'E%' THEN 1 ELSE 0 END) AS IFE, "
            "SUM(CASE WHEN df1.Rev LIKE 'P%' THEN 1 ELSE 0 END) AS IFP, "
            "SUM(CASE WHEN df1.Rev LIKE 'C%' THEN 1 ELSE 0 END) AS AFC "
            "FROM df df1 "
            "GROUP BY df1.Name, df1.Title"
        )
        try:
            return sqldf(query, {"df": df})
        except Exception as e:
            messagebox.showwarning("Aggregation Failed", f"{e}")
            return df

    def apply_filter(self):
        if self.df_raw.empty:
            messagebox.showinfo("Info", "No data to filter.")
            return

        filtered_df = self.df_raw.copy()

        if self.selected_uboc_values and "ALL" not in self.selected_uboc_values:
            mask = pd.Series(False, index=filtered_df.index)

            if "BLANK" in self.selected_uboc_values:
                mask |= (
                    filtered_df["UBOC Return Code"].isna() |
                    (filtered_df["UBOC Return Code"].str.strip() == "")
                )

            normal_values = [v for v in self.selected_uboc_values if v != "BLANK"]
            if normal_values:
                mask |= filtered_df["UBOC Return Code"].isin(normal_values)

            filtered_df = filtered_df[mask]

        date_col = "Workflow Start"
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
            except Exception as e:
                messagebox.showerror("Date Filter Error", str(e))
                return

        if filtered_df.empty:
            messagebox.showinfo("Info", "No matching records found.")
            self.table.model.df = pd.DataFrame()
        else:
            agg_df = self.apply_aggregation(filtered_df)
            self.table.model.df = agg_df

        self.table.redraw()

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df_raw_original = pd.read_excel(file_path, dtype=str)
            self.current_file = file_path
            self.df_raw = self.apply_initial_base_sql(self.df_raw_original)
            self.apply_filter()
            self._update_status(f"Loaded and filtered: {os.path.basename(file_path)}")
            self.refresh_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            self.df_raw_original = pd.DataFrame()

    def refresh_file(self):
        if not self.current_file:
            return
        try:
            self.df_raw_original = pd.read_excel(self.current_file, dtype=str)
            self.df_raw = self.apply_initial_base_sql(self.df_raw_original)
            self.apply_filter()
            self._update_status(f"Refreshed: {os.path.basename(self.current_file)}")
        except Exception as e:
            messagebox.showerror("Refresh Error", str(e))

    def save_file(self):
        df_to_save = self.table.model.df
        if df_to_save.empty:
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
            df_to_save.to_excel(file_path, index=False)
            self._update_status(f"Data saved to: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def _update_status(self, message: str):
        self.file_label.config(text=message)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTableApp(root)
    root.mainloop()
