#!/usr/bin/env python
# coding: utf-8

# In[11]:


import os
import pandas as pd
import numpy as np
from scipy.stats import skew, kurtosis
from tkinter import filedialog, messagebox, BOTH, RIGHT, LEFT, Y, W
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk


class ExcelSummaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CoreSummaryStat")
        self.root.geometry("1100x700")

        # Set custom icon if available
        icon_path = "myicon.ico"  # replace with your custom icon
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        # Sidebar
        sidebar = ttk.Frame(root, padding=10)
        sidebar.pack(side=LEFT, fill=Y)

        # Logo
        logo_path = "myicon.png"  # replace with your logo if available
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path).resize((120, 120))
                self.logo_img = ImageTk.PhotoImage(img)
                logo_label = ttk.Label(sidebar, image=self.logo_img)
                logo_label.pack(pady=10)
            except Exception:
                pass

        # File selection
        ttk.Button(sidebar, text="Select Excel Files", command=self.load_files, bootstyle=PRIMARY).pack(fill=X, pady=5)

        # Measure selection
        self.measures = [
            "Mean", "Median", "Mode", "Range", "Variance", "Standard Deviation",
            "Skewness", "Kurtosis", "Max", "Min", "Coefficient of Variation",
            "Q1", "Q2", "Q3", "Quartile Deviation", "Correlation"
        ]
        self.measure_vars = {m: ttk.BooleanVar(value=True) for m in self.measures}

        ttk.Label(sidebar, text="Select Measures", font=("Segoe UI", 10, "bold")).pack(pady=5)
        measure_frame = ttk.Frame(sidebar)
        measure_frame.pack(fill=Y, expand=True)

        for m in self.measures:
            ttk.Checkbutton(measure_frame, text=m, variable=self.measure_vars[m]).pack(anchor=W)

        # Column selection (dynamic later)
        ttk.Label(sidebar, text="Select Columns", font=("Segoe UI", 10, "bold")).pack(pady=5)
        self.column_frame = ttk.Frame(sidebar)
        self.column_frame.pack(fill=Y, expand=True, pady=5)
        self.column_vars = {}

        # Action buttons
        ttk.Button(sidebar, text="Generate Summary", command=self.generate_summary, bootstyle=SUCCESS).pack(fill=X, pady=5)
        ttk.Button(sidebar, text="Save Output", command=self.save_output, bootstyle=INFO).pack(fill=X, pady=5)

        # Preview area (Treeview)
        self.preview = ttk.Treeview(root, columns=("Column", "Measure", "Value"), show="headings", height=25)
        self.preview.heading("Column", text="Column")
        self.preview.heading("Measure", text="Measure")
        self.preview.heading("Value", text="Value")
        self.preview.column("Column", width=120, anchor=W)
        self.preview.column("Measure", width=150, anchor=W)
        self.preview.column("Value", width=150, anchor=W)
        self.preview.pack(side=RIGHT, fill=BOTH, expand=True, padx=10, pady=10)

        self.files = []
        self.summary = None

    def load_files(self):
        self.files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not self.files:
            return

        # Preview first file’s columns for selection
        df = pd.read_excel(self.files[0], sheet_name=None)
        sample_sheet = list(df.values())[0]

        # Clear previous column checkboxes
        for widget in self.column_frame.winfo_children():
            widget.destroy()

        self.column_vars = {}
        for col in sample_sheet.columns:
            var = ttk.BooleanVar(value=True)
            self.column_vars[col] = var
            ttk.Checkbutton(self.column_frame, text=col, variable=var).pack(anchor=W)

    def generate_summary(self):
        if not self.files:
            messagebox.showerror("Error", "No files selected")
            return

        summary_data = {}
        all_stats = []

        for file in self.files:
            sheets = pd.read_excel(file, sheet_name=None)
            for sheet_name, df in sheets.items():
                numeric_df = df.select_dtypes(include=[np.number])

                # filter columns
                selected_cols = [col for col, var in self.column_vars.items() if var.get()]
                if selected_cols:
                    numeric_df = numeric_df[selected_cols]

                sheet_summary = {}
                for col in numeric_df.columns:
                    stats = {}
                    data = numeric_df[col].dropna()

                    if self.measure_vars["Mean"].get():
                        stats["Mean"] = data.mean()
                    if self.measure_vars["Median"].get():
                        stats["Median"] = data.median()
                    if self.measure_vars["Mode"].get():
                        stats["Mode"] = data.mode().iloc[0] if not data.mode().empty else np.nan
                    if self.measure_vars["Range"].get():
                        stats["Range"] = data.max() - data.min()
                    if self.measure_vars["Variance"].get():
                        stats["Variance"] = data.var()
                    if self.measure_vars["Standard Deviation"].get():
                        stats["Standard Deviation"] = data.std()
                    if self.measure_vars["Skewness"].get():
                        stats["Skewness"] = skew(data)
                    if self.measure_vars["Kurtosis"].get():
                        stats["Kurtosis"] = kurtosis(data)
                    if self.measure_vars["Max"].get():
                        stats["Max"] = data.max()
                    if self.measure_vars["Min"].get():
                        stats["Min"] = data.min()
                    if self.measure_vars["Coefficient of Variation"].get():
                        stats["CV"] = data.std() / data.mean() if data.mean() != 0 else np.nan
                    if self.measure_vars["Q1"].get():
                        stats["Q1"] = data.quantile(0.25)
                    if self.measure_vars["Q2"].get():
                        stats["Q2"] = data.quantile(0.5)
                    if self.measure_vars["Q3"].get():
                        stats["Q3"] = data.quantile(0.75)
                    if self.measure_vars["Quartile Deviation"].get():
                        stats["QD"] = (data.quantile(0.75) - data.quantile(0.25)) / 2

                    sheet_summary[col] = stats

                    # Collect for master descriptive sheet
                    stats_copy = stats.copy()
                    stats_copy["File"] = os.path.basename(file)
                    stats_copy["Sheet"] = sheet_name
                    stats_copy["Column"] = col
                    all_stats.append(stats_copy)

                # correlation
                if self.measure_vars["Correlation"].get() and not numeric_df.empty:
                    sheet_summary["Correlation"] = numeric_df.corr()

                summary_data[f"{os.path.basename(file)} - {sheet_name}"] = sheet_summary

        self.summary = summary_data
        self.master_descriptive = pd.DataFrame(all_stats)

        # Update preview
        for i in self.preview.get_children():
            self.preview.delete(i)

        for sheet, cols in summary_data.items():
            for col, stats in cols.items():
                if col == "Correlation":
                    self.preview.insert("", "end", values=(f"{sheet} Correlation", "", "See Excel Output"))
                else:
                    for measure, value in stats.items():
                        self.preview.insert("", "end", values=(col, measure, round(value, 4) if pd.notna(value) else "NA"))

    def save_output(self):
        if not self.summary:
            messagebox.showerror("Error", "No summary generated")
            return

        out_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
        if not out_file:
            return

        with pd.ExcelWriter(out_file) as writer:
            for sheet, cols in self.summary.items():
                dfs = []
                for col, stats in cols.items():
                    if col == "Correlation":
                        stats.to_excel(writer, sheet_name=f"{sheet}_Correlation")
                    else:
                        df = pd.DataFrame(stats, index=[col])
                        dfs.append(df)
                if dfs:
                    pd.concat(dfs).to_excel(writer, sheet_name=sheet)

            # Add Master Descriptive Sheet
            self.master_descriptive.to_excel(writer, sheet_name="Master_Descriptive", index=False)

        messagebox.showinfo("Success", f"Summary saved to {out_file}")


if __name__ == "__main__":
    app = ttk.Window(themename="flatly")  # try "darkly", "cosmo" too
    ExcelSummaryApp(app)
    app.mainloop()


# In[ ]:




