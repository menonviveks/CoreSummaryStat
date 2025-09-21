#!/usr/bin/env python
# coding: utf-8

# In[7]:


import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+ DPI awareness
except:
    pass

class ExcelSummaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CoreSummaryStat")
        root.iconbitmap("monitoring.ico")
        self.filepath = None
        self.df_dict = None  # Store each sheet separately
        self.sheetnames = []

        # Notebook with two tabs
        self.notebook = ttk.Notebook(root)
        self.measures_tab = ttk.Frame(self.notebook)
        self.plots_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.measures_tab, text="Measures")
        self.notebook.add(self.plots_tab, text="Plots")
        self.notebook.pack(fill="both", expand=True)

        self.setup_measures_tab()
        self.setup_plots_tab()

    # ================== Measures Tab ===================
    def setup_measures_tab(self):
        sidebar = ttk.Frame(self.measures_tab, width=200)
        sidebar.pack(side="left", fill="y")

        ttk.Button(sidebar, text="Load File", command=self.load_file).pack(fill="x", pady=5)

        ttk.Label(sidebar, text="Select Sheet").pack(pady=5)
        self.sheet_combo_m = ttk.Combobox(sidebar, state="readonly")
        self.sheet_combo_m.bind("<<ComboboxSelected>>", self.load_columns_measures)
        self.sheet_combo_m.pack(fill="x")

        ttk.Label(sidebar, text="Select Columns").pack(pady=5)
        self.col_listbox_m = tk.Listbox(sidebar, selectmode="multiple", height=10)
        self.col_listbox_m.pack(fill="x")

        # Measures
        self.measures = {
            "Mean": np.mean,
            "Median": np.median,
            "Mode": lambda x: stats.mode(x, nan_policy="omit")[0][0] if len(x) > 0 else np.nan,
            "Std Dev": np.std,
            "Variance": np.var,
            "Min": np.min,
            "Max": np.max,
            "Range": lambda x: np.max(x) - np.min(x),
            "Q1": lambda x: np.percentile(x, 25),
            "Q3": lambda x: np.percentile(x, 75),
            "IQR": lambda x: np.percentile(x, 75) - np.percentile(x, 25),
            "Quartile Deviation": lambda x: (np.percentile(x, 75) - np.percentile(x, 25)) / 2,
            "Mean Absolute Deviation": lambda x: np.mean(np.abs(x - np.mean(x))),
            "Coefficient of Variation": lambda x: (np.std(x) / np.mean(x)) * 100 if np.mean(x) != 0 else np.nan,
        }
        self.measure_vars = {}
        for m in self.measures:
            var = tk.BooleanVar()
            ttk.Checkbutton(sidebar, text=m, variable=var).pack(anchor="w")
            self.measure_vars[m] = var

        ttk.Button(sidebar, text="Preview Measures", command=self.preview_measures).pack(fill="x", pady=5)
        ttk.Button(sidebar, text="Save Measures (Excel)", command=self.save_measures).pack(fill="x", pady=5)

        # Scrollable preview
        self.measure_preview = tk.Text(self.measures_tab, wrap="word")
        self.measure_preview.pack(side="right", fill="both", expand=True)
        scroll_m = ttk.Scrollbar(self.measure_preview, command=self.measure_preview.yview)
        self.measure_preview.config(yscrollcommand=scroll_m.set)
        scroll_m.pack(side="right", fill="y")

    # ================== Plots Tab ===================
    def setup_plots_tab(self):
        sidebar = ttk.Frame(self.plots_tab, width=200)
        sidebar.pack(side="left", fill="y")

        ttk.Label(sidebar, text="Select Sheet").pack(pady=5)
        self.sheet_combo_p = ttk.Combobox(sidebar, state="readonly")
        self.sheet_combo_p.bind("<<ComboboxSelected>>", self.load_columns_plots)
        self.sheet_combo_p.pack(fill="x")

        ttk.Label(sidebar, text="Select Columns").pack(pady=5)
        self.col_listbox_p = tk.Listbox(sidebar, selectmode="multiple", height=10)
        self.col_listbox_p.pack(fill="x")

        ttk.Label(sidebar, text="Select Plot Types").pack(pady=5)
        self.plot_types = ["Histogram", "Boxplot", "Violinplot", "Density Plot", "Correlation Heatmap"]
        self.plot_vars = {}
        for pt in self.plot_types:
            var = tk.BooleanVar()
            ttk.Checkbutton(sidebar, text=pt, variable=var).pack(anchor="w")
            self.plot_vars[pt] = var

        ttk.Button(sidebar, text="Preview Plots", command=self.preview_plots).pack(fill="x", pady=5)
        ttk.Button(sidebar, text="Save Plots (PNG)", command=self.save_plots).pack(fill="x", pady=5)

        # Scrollable canvas for plot preview
        self.plot_canvas_frame = tk.Frame(self.plots_tab)
        self.plot_canvas_frame.pack(side="right", fill="both", expand=True)
        self.plot_canvas = tk.Canvas(self.plot_canvas_frame)
        self.scroll_y = ttk.Scrollbar(self.plot_canvas_frame, orient="vertical", command=self.plot_canvas.yview)
        self.scroll_y.pack(side="right", fill="y")
        self.plot_canvas.pack(side="left", fill="both", expand=True)
        self.plot_canvas.configure(yscrollcommand=self.scroll_y.set)
        self.plot_frame = tk.Frame(self.plot_canvas)
        self.plot_canvas.create_window((0,0), window=self.plot_frame, anchor="nw")
        self.plot_frame.bind("<Configure>", lambda e: self.plot_canvas.configure(scrollregion=self.plot_canvas.bbox("all")))

    # ================== File Loading ===================
    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not self.filepath: return

        if self.filepath.endswith(".csv"):
            df = pd.read_csv(self.filepath)
            self.df_dict = {"CSV": df}
            self.sheetnames = ["CSV"]
        else:
            xls = pd.ExcelFile(self.filepath)
            self.sheetnames = ["Multiple Sheets"] + xls.sheet_names
            self.df_dict = {}  # Will fill when sheets are selected

        # Update both sheet dropdowns
        self.sheet_combo_m["values"] = self.sheetnames
        self.sheet_combo_p["values"] = self.sheetnames
        self.sheet_combo_m.current(0)
        self.sheet_combo_p.current(0)
        self.load_columns_measures()
        self.load_columns_plots()

    # ============ Columns Loading ===================
    def load_columns_measures(self, event=None):
        self.col_listbox_m.delete(0, "end")
        sheet = self.sheet_combo_m.get()
        if sheet == "Multiple Sheets":
            self.ask_sheets_selection()
            # Show first selected sheet columns
            first_sheet = self.df_dict[list(self.df_dict.keys())[0]]
            for col in first_sheet.select_dtypes(include=[np.number]).columns:
                self.col_listbox_m.insert("end", col)
        else:
            if sheet not in self.df_dict:
                self.df_dict[sheet] = pd.read_excel(self.filepath, sheet_name=sheet)
            df = self.df_dict[sheet]
            for col in df.select_dtypes(include=[np.number]).columns:
                self.col_listbox_m.insert("end", col)

    def load_columns_plots(self, event=None):
        self.col_listbox_p.delete(0, "end")
        sheet = self.sheet_combo_p.get()
        if sheet == "Multiple Sheets":
            self.ask_sheets_selection()
            first_sheet = self.df_dict[list(self.df_dict.keys())[0]]
            for col in first_sheet.select_dtypes(include=[np.number]).columns:
                self.col_listbox_p.insert("end", col)
        else:
            if sheet not in self.df_dict:
                self.df_dict[sheet] = pd.read_excel(self.filepath, sheet_name=sheet)
            df = self.df_dict[sheet]
            for col in df.select_dtypes(include=[np.number]).columns:
                self.col_listbox_p.insert("end", col)

    # ================== Multiple Sheets Selection ===================
    def ask_sheets_selection(self):
        xls = pd.ExcelFile(self.filepath)
        sheet_list = xls.sheet_names

        popup = tk.Toplevel(self.root)
        popup.title("Select Sheets")
        popup.geometry("250x300")
        vars_dict = {}

        tk.Label(popup, text="Select sheets to include:").pack(pady=5)
        frame = tk.Frame(popup)
        frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for sheet in sheet_list:
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(scrollable_frame, text=sheet, variable=var).pack(anchor="w")
            vars_dict[sheet] = var

        def submit_selection():
            selected_sheets = [s for s, v in vars_dict.items() if v.get()]
            if not selected_sheets:
                messagebox.showerror("Error", "No sheets selected!")
                return
            self.df_dict = {s: pd.read_excel(self.filepath, sheet_name=s) for s in selected_sheets}
            popup.destroy()

        ttk.Button(popup, text="OK", command=submit_selection).pack(pady=5)
        popup.transient(self.root)
        popup.grab_set()
        self.root.wait_window(popup)

    # ================== Measures Preview/Save ===================
    def preview_measures(self):
        selected_cols = [self.col_listbox_m.get(i) for i in self.col_listbox_m.curselection()]
        selected_measures = [m for m, var in self.measure_vars.items() if var.get()]
        self.measure_preview.delete(1.0, "end")
        if not selected_cols or not selected_measures:
            self.measure_preview.insert("end", "No columns or measures selected.")
            return

        for sheet, df in self.df_dict.items():
            self.measure_preview.insert("end", f"\n--- Sheet: {sheet} ---\n")
            for col in selected_cols:
                if col not in df.columns:
                    continue
                data = df[col].dropna().values
                if len(data) == 0:
                    self.measure_preview.insert("end", f"Column: {col} - No numeric data.\n")
                    continue
                self.measure_preview.insert("end", f"\nColumn: {col}\n")
                for m in selected_measures:
                    try:
                        val = self.measures[m](data)
                        self.measure_preview.insert("end", f"  {m}: {val:.4f}\n")
                    except:
                        self.measure_preview.insert("end", f"  {m}: Error\n")

    def save_measures(self):
        selected_cols = [self.col_listbox_m.get(i) for i in self.col_listbox_m.curselection()]
        selected_measures = [m for m, var in self.measure_vars.items() if var.get()]
        if not selected_cols or not selected_measures:
            messagebox.showerror("Error", "No columns or measures selected.")
            return

        results = []
        for sheet, df in self.df_dict.items():
            for col in selected_cols:
                if col not in df.columns:
                    continue
                data = df[col].dropna().values
                row = {"Sheet": sheet, "Column": col}
                for m in selected_measures:
                    try:
                        row[m] = self.measures[m](data)
                    except:
                        row[m] = np.nan
                results.append(row)

        out_df = pd.DataFrame(results)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files","*.xlsx")])
        if save_path:
            out_df.to_excel(save_path, index=False)
            messagebox.showinfo("Saved", f"Measures saved to {save_path}")

    # ================== Plots Preview/Save ===================
    def preview_plots(self):
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
        selected_cols = [self.col_listbox_p.get(i) for i in self.col_listbox_p.curselection()]
        selected_plots = [pt for pt, var in self.plot_vars.items() if var.get()]
        if len(selected_cols) == 0 or len(selected_plots) == 0:
            messagebox.showerror("Error", "No columns or plots selected.")
            return

        pastel_colors = sns.color_palette("pastel")

        for sheet, df in self.df_dict.items():
            for pt in selected_plots:
                if pt in ["Boxplot", "Violinplot"]:
                    for i, col in enumerate(selected_cols):
                        if col not in df.columns:
                            continue
                        data = df[col].dropna()
                        if len(data) == 0:
                            continue
                        fig, ax = plt.subplots(figsize=(6,4))
                        color = pastel_colors[i % len(pastel_colors)]
                        if pt == "Boxplot":
                            sns.boxplot(y=data, color=color, ax=ax)
                        else:
                            sns.violinplot(y=data, color=color, ax=ax)
                        ax.set_title(f"{pt} – {col} (Sheet: {sheet})")
                        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
                        canvas.draw()
                        canvas.get_tk_widget().pack(fill="both", expand=True, pady=10)
                        plt.close(fig)
                elif pt in ["Histogram", "Density Plot"]:
                    fig, ax = plt.subplots(figsize=(6, max(4, len(selected_cols)*0.6)))
                    for i, col in enumerate(selected_cols):
                        if col not in df.columns:
                            continue
                        data = df[col].dropna()
                        if len(data) == 0:
                            continue
                        color = pastel_colors[i % len(pastel_colors)]
                        if pt == "Histogram":
                            sns.histplot(data, bins=20, alpha=0.7, color=color, label=col, ax=ax)
                        else:
                            sns.kdeplot(data, fill=True, alpha=0.5, color=color, label=col, ax=ax)
                    ax.set_title(f"{pt} – Sheet: {sheet}")
                    ax.legend()
                    canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
                    canvas.draw()
                    canvas.get_tk_widget().pack(fill="both", expand=True, pady=10)
                    plt.close(fig)
                elif pt == "Correlation Heatmap":
                    if len(selected_cols) < 2:
                        continue
                    fig, ax = plt.subplots(figsize=(6, max(4, len(selected_cols)*0.6)))
                    corr = df[selected_cols].corr()
                    mask = np.triu(np.ones_like(corr, dtype=bool))
                    sns.heatmap(corr, mask=mask, cmap="RdYlGn", annot=True, fmt=".2f", ax=ax)
                    ax.set_title(f"Correlation Heatmap – Sheet: {sheet}")
                    canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
                    canvas.draw()
                    canvas.get_tk_widget().pack(fill="both", expand=True, pady=10)
                    plt.close(fig)

    def save_plots(self):
        selected_cols = [self.col_listbox_p.get(i) for i in self.col_listbox_p.curselection()]
        selected_plots = [pt for pt, var in self.plot_vars.items() if var.get()]
        if len(selected_cols) == 0 or len(selected_plots) == 0:
            messagebox.showerror("Error", "No columns or plots selected.")
            return

        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        pastel_colors = sns.color_palette("pastel")

        for sheet, df in self.df_dict.items():
            for pt in selected_plots:
                if pt in ["Boxplot", "Violinplot"]:
                    for i, col in enumerate(selected_cols):
                        if col not in df.columns:
                            continue
                        data = df[col].dropna()
                        if len(data) == 0:
                            continue
                        fig, ax = plt.subplots(figsize=(6,4))
                        color = pastel_colors[i % len(pastel_colors)]
                        if pt == "Boxplot":
                            sns.boxplot(y=data, color=color, ax=ax)
                        else:
                            sns.violinplot(y=data, color=color, ax=ax)
                        ax.set_title(f"{pt} – {col} (Sheet: {sheet})")
                        filename = f"{pt}_{col}_{sheet}.png"
                        fig.savefig(os.path.join(folder_path, filename))
                        plt.close(fig)
                elif pt in ["Histogram", "Density Plot"]:
                    fig, ax = plt.subplots(figsize=(6, max(4, len(selected_cols)*0.6)))
                    for i, col in enumerate(selected_cols):
                        if col not in df.columns:
                            continue
                        data = df[col].dropna()
                        if len(data) == 0:
                            continue
                        color = pastel_colors[i % len(pastel_colors)]
                        if pt == "Histogram":
                            sns.histplot(data, bins=20, alpha=0.7, color=color, label=col, ax=ax)
                        else:
                            sns.kdeplot(data, fill=True, alpha=0.5, color=color, label=col, ax=ax)
                    ax.set_title(f"{pt} – Sheet: {sheet}")
                    ax.legend()
                    filename = f"{pt}_{sheet}.png"
                    fig.savefig(os.path.join(folder_path, filename))
                    plt.close(fig)
                elif pt == "Correlation Heatmap":
                    if len(selected_cols) < 2:
                        continue
                    fig, ax = plt.subplots(figsize=(6, max(4, len(selected_cols)*0.6)))
                    corr = df[selected_cols].corr()
                    mask = np.triu(np.ones_like(corr, dtype=bool))
                    sns.heatmap(corr, mask=mask, cmap="RdYlGn", annot=True, fmt=".2f", ax=ax)
                    ax.set_title(f"Correlation Heatmap – Sheet: {sheet}")
                    filename = f"Correlation_Heatmap_{sheet}.png"
                    fig.savefig(os.path.join(folder_path, filename))
                    plt.close(fig)

        messagebox.showinfo("Saved", f"All plots saved as PNGs in {folder_path}")


# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSummaryApp(root)
    root.mainloop()


# In[ ]:




