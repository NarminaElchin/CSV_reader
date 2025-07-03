import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import os
import time


class AdaptiveFiltersFrame(tk.Frame):
    def __init__(self, parent, max_widget_width=200, **kwargs):
        super().__init__(parent, **kwargs)
        self.max_widget_width = max_widget_width
        self.filter_widgets = []
        self.parent = parent
        self.bind('<Configure>', self.on_resize)

    def add_filter_widget(self, widget):
        self.filter_widgets.append(widget)
        widget.master = self  # set master to this frame
        widget.grid_configure(padx=5, pady=5, sticky="ew")

    def on_resize(self, event):
        width = event.width
        max_per_row = max(1, width // self.max_widget_width)

        for w in self.filter_widgets:
            w.grid_forget()

        for idx, widget in enumerate(self.filter_widgets):
            row = idx // max_per_row
            col = idx % max_per_row
            widget.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        for col in range(max_per_row):
            self.grid_columnconfigure(col, weight=1)

class CSVTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Data Viewer & Excel Exporter")
        self.root.state('zoomed')

        self.df_raw_original = None
        self.df_raw = None
        self.current_file = None
        self.selected_uboc_values = []
        self.table = None

        self._setup_widgets()
        self._setup_filters()
        self._initialize_table_frame()

    def _setup_widgets(self):
        self.file_label = tk.Label(self.root, text="No file loaded", pady=10, font=("Arial", 12))
        self.file_label.pack()

        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=5)

        tk.Button(self.button_frame, text="Load CSV File", command=self.load_file).pack(side=tk.LEFT, padx=5)
        self.refresh_button = tk.Button(self.button_frame, text="Refresh", command=self.refresh_file, state=tk.DISABLED)
        self.refresh_button.pack(side=tk.LEFT, padx=5)
        self.save_button = tk.Button(self.button_frame, text="Save to Excel", command=self.save_file, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

    def _setup_filters(self):
        # Create adaptive filter container
        self.filter_frame = AdaptiveFiltersFrame(self.root, max_widget_width=250)
        self.filter_frame.pack(fill=tk.X, pady=10, padx=10)

        uboc_frame = tk.Frame(self.filter_frame)
        uboc_label = tk.Label(uboc_frame, text="UBOC Code:")
        uboc_label.pack(side=tk.LEFT, padx=(0, 5))
        self.uboc_button_var = tk.StringVar(value="Select UBOC Code ▼")
        self.uboc_dropdown_button = tk.Button(uboc_frame, textvariable=self.uboc_button_var, width=25,
                                              command=self.show_uboc_dropdown)
        self.uboc_dropdown_button.pack(side=tk.LEFT)
        self.filter_frame.add_filter_widget(uboc_frame)

        # Date From filter
        date_from_frame = tk.Frame(self.filter_frame)
        date_from_label = tk.Label(date_from_frame, text="Workflow Start From:")
        date_from_label.pack(side=tk.LEFT, padx=(0, 5))
        self.date_from = DateEntry(date_from_frame, date_pattern='dd/mm/yyyy', width=12)
        self.date_from.pack(side=tk.LEFT)
        self.date_from.delete(0, tk.END)
        self.filter_frame.add_filter_widget(date_from_frame)

        # Date To filter
        date_to_frame = tk.Frame(self.filter_frame)
        date_to_label = tk.Label(date_to_frame, text="To:")
        date_to_label.pack(side=tk.LEFT, padx=(0, 5))
        self.date_to = DateEntry(date_to_frame, date_pattern='dd/mm/yyyy', width=12)
        self.date_to.pack(side=tk.LEFT)
        self.date_to.delete(0, tk.END)
        self.filter_frame.add_filter_widget(date_to_frame)

        # Apply Filters button
        self.apply_button = tk.Button(self.filter_frame, text="Apply Filters", command=self.apply_filter)
        self.filter_frame.add_filter_widget(self.apply_button)

    def show_uboc_dropdown(self):
        if self.df_raw is None or "UBOC Return Code" not in self.df_raw.columns:
            return

        top = tk.Toplevel(self.root)
        top.title("Select UBOC Code")
        top.geometry("+%d+%d" % (self.root.winfo_pointerx(), self.root.winfo_pointery()))
        top.grab_set()

        self.uboc_vars = {}
        all_values = sorted(self.df_raw["UBOC Return Code"].dropna().unique())
        all_values = all_values + ["BLANK"]

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
                display_text += f"... (+{len(selected) - 3})"
            self.uboc_button_var.set(display_text or "Select UBOC Code ▼")
            top.destroy()

        tk.Button(top, text="Confirm", command=confirm_selection).pack(pady=5)

    def _initialize_table_frame(self):
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)

    def _initialize_table(self, df):
        from pandastable import Table
        if self.table is None:
            self.table = Table(self.table_frame, dataframe=df, showtoolbar=False, showstatusbar=True)
            self.table.show()
        else:
            self.table.model.df = df
            self.table.redraw()

    def apply_initial_base_sql(self, df):
        try:
            mask_no_v01_x01 = ~df['Name'].isin(
                df[df['Rev'].isin(['V01', 'X01'])]['Name']
            )
            mask_rev = df['Rev'] != 'A01'
            mask_subject = df['Subject'].isin(['EXECUTE', 'DEFINE/EXECUTE'])
            return df[mask_no_v01_x01 & mask_rev & mask_subject]
        except Exception as e:
            messagebox.showerror("Filter Error", f"Base filtering failed: {e}")
            return df

    def apply_aggregation(self, df):
        import pandas as pd
        try:
            agg_df = df.copy()
            agg_df['Rev'] = agg_df['Rev'].fillna('')
            agg_df['IFR'] = (agg_df['Rev'].str.startswith('B')).fillna(False).astype(int)
            agg_df['AFD'] = (agg_df['Rev'].str.startswith('D')).fillna(False).astype(int)
            agg_df['AFU'] = (agg_df['Rev'].str.startswith('U')).fillna(False).astype(int)
            agg_df['AFH'] = (agg_df['Rev'].str.startswith('H')).fillna(False).astype(int)
            agg_df['IFI'] = (agg_df['Rev'].str.startswith('I')).fillna(False).astype(int)
            agg_df['IFE'] = (agg_df['Rev'].str.startswith('E')).fillna(False).astype(int)
            agg_df['IFP'] = (agg_df['Rev'].str.startswith('P')).fillna(False).astype(int)
            agg_df['AFC'] = (agg_df['Rev'].str.startswith('C')).fillna(False).astype(int)
            return agg_df.groupby(['Name', 'Title'])[
                ['IFR', 'AFD', 'AFU', 'AFH', 'IFI', 'IFE', 'IFP', 'AFC']].sum().reset_index()
        except Exception as e:
            messagebox.showwarning("Aggregation Failed", f"{e}")
            return df

    def apply_filter(self):
        import pandas as pd
        if self.df_raw is None or self.df_raw.empty:
            messagebox.showinfo("Info", "No data to filter.")
            if self.table:
                self._initialize_table(pd.DataFrame())
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
            self._initialize_table(pd.DataFrame())
        else:
            agg_df = self.apply_aggregation(filtered_df)
            self._initialize_table(agg_df)

    def load_file(self):
        import pandas as pd
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            if not os.path.isfile(file_path) or not file_path.lower().endswith('.csv'):
                raise ValueError("Selected file is not a valid CSV file.")

            self.df_raw_original = pd.read_csv(file_path, dtype=str, encoding='utf-8', on_bad_lines='warn')

            required_columns = ['Name', 'Title', 'Rev', 'Subject']
            missing_columns = [col for col in required_columns if col not in self.df_raw_original.columns]
            if missing_columns:
                raise ValueError(f"CSV missing required columns: {', '.join(missing_columns)}")

            if self.df_raw_original['Rev'].isna().mean() > 0.5:
                messagebox.showwarning("Data Warning",
                                       "Over 50% of 'Rev' values are missing, which may affect results.")

            self.current_file = file_path
            self.df_raw = self.apply_initial_base_sql(self.df_raw_original)
            self.apply_filter()
            self._update_status(f"Loaded and filtered: {os.path.basename(file_path)}")
            self.refresh_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load CSV: {e}")
            self.df_raw_original = None
            self.df_raw = None

    def refresh_file(self):
        import pandas as pd
        if not self.current_file:
            return
        try:
            self.df_raw_original = pd.read_csv(self.current_file, dtype=str, encoding='utf-8', on_bad_lines='warn')

            required_columns = ['Name', 'Title', 'Rev', 'Subject']
            missing_columns = [col for col in required_columns if col not in self.df_raw_original.columns]
            if missing_columns:
                raise ValueError(f"CSV missing required columns: {', '.join(missing_columns)}")

            if self.df_raw_original['Rev'].isna().mean() > 0.5:
                messagebox.showwarning("Data Warning",
                                       "Over 50% of 'Rev' values are missing, which may affect results.")

            self.df_raw = self.apply_initial_base_sql(self.df_raw_original)

            self.selected_uboc_values = []
            self.uboc_button_var.set("Select UBOC Code ▼")
            self.date_from.delete(0, tk.END)
            self.date_to.delete(0, tk.END)

            agg_df = self.apply_aggregation(self.df_raw)
            self._initialize_table(agg_df)
            self._update_status(f"Refreshed: {os.path.basename(self.current_file)}")
        except Exception as e:
            messagebox.showerror("Refresh Error", f"Failed to refresh CSV: {e}")

    def save_file(self):
        import pandas as pd
        import openpyxl
        if self.table is None or self.table.model.df.empty:
            messagebox.showerror("Error", "No data to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Aggregated Data to Excel"
        )
        if not file_path:
            return

        try:
            self.table.model.df.to_excel(file_path, index=False, engine='openpyxl')
            self._update_status(f"Data saved to: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save Excel file: {e}")

    def _update_status(self, message: str):
        self.file_label.config(text=message)


if __name__ == "__main__":
    start_time = time.time()
    root = tk.Tk()
    splash = tk.Label(root, text="Loading CSV Data Viewer...", font=("Arial", 14))
    splash.pack(pady=20)
    root.update()
    app = CSVTableApp(root)
    splash.destroy()
    print(f"Startup time: {time.time() - start_time:.2f} seconds")
    root.mainloop()