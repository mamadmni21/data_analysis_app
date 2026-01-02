import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt


class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Viewer")
        self.root.geometry("1200x700")

        self.excel_file_path = None
        self.df = None
        self.filtered_df = None

        self.main_frame = tk.Frame(root)
        self.analysis_frame = tk.Frame(root)
        self.pie_chart_frame = tk.Frame(root)
        self.salary_frame = tk.Frame(root)

        self.result_col_vars = []

        self.create_main_page()
        self.create_analysis_page()
        self.create_pie_chart_page()
        self.create_salary_page()

        self.show_main_page()

    def create_main_page(self):
        frame = tk.Frame(self.main_frame)
        frame.pack(pady=10)

        tk.Button(frame, text="Browse File", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Refresh Data", command=self.load_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Analyze", command=self.show_analysis_page).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Pie Chart", command=self.show_pie_chart_page).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Salary", command=self.show_salary_page).pack(side=tk.LEFT, padx=5)

        self.file_label = tk.Label(frame, text="No file selected", fg="blue")
        self.file_label.pack(side=tk.LEFT, padx=10)

        table_frame = tk.Frame(self.main_frame)
        table_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        self.tree = ttk.Treeview(table_frame, show="headings")
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(expand=True, fill=tk.BOTH)

    def create_analysis_page(self):
        top_frame = tk.Frame(self.analysis_frame)
        top_frame.pack(pady=10)

        tk.Label(top_frame, text="Data Analysis", font=("Arial", 14, "bold")).pack()

        control_frame = tk.Frame(self.analysis_frame)
        control_frame.pack(pady=5)

        self.column_var = tk.StringVar()
        self.column_dropdown = ttk.Combobox(control_frame, textvariable=self.column_var, state="readonly")
        self.column_dropdown.pack(side=tk.LEFT, padx=5)

        self.filter_entry = tk.Entry(control_frame)
        self.filter_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(control_frame, text="Filter", command=self.filter_data).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Total", command=self.calculate_total).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Back", command=self.show_main_page).pack(side=tk.LEFT, padx=5)

        self.filter_note = tk.Label(self.analysis_frame, text="Type value or 'ALL'", fg="gray")
        self.filter_note.pack()

        self.analysis_display_frame = tk.Frame(self.analysis_frame)
        self.analysis_display_frame.pack(expand=True, fill=tk.BOTH)

        self.analysis_tree = ttk.Treeview(self.analysis_display_frame, show="headings")
        self.analysis_tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)

    def create_pie_chart_page(self):
        control_frame = tk.Frame(self.pie_chart_frame)
        control_frame.pack(pady=10)

        tk.Label(control_frame, text="Pie Chart Generator", font=("Arial", 14, "bold")).grid(row=0, columnspan=2, pady=5)

        tk.Label(control_frame, text="Label Column 1").grid(row=1, column=0)
        tk.Label(control_frame, text="Label Column 2").grid(row=2, column=0)

        self.col1_var = tk.StringVar()
        self.col2_var = tk.StringVar()

        self.col1_dropdown = ttk.Combobox(control_frame, textvariable=self.col1_var, state="readonly")
        self.col2_dropdown = ttk.Combobox(control_frame, textvariable=self.col2_var, state="readonly")

        self.col1_dropdown.grid(row=1, column=1, padx=5, pady=2)
        self.col2_dropdown.grid(row=2, column=1, padx=5, pady=2)

        tk.Button(control_frame, text="Show Pie Chart", command=self.generate_pie_chart).grid(row=4, columnspan=2, pady=10)
        tk.Button(control_frame, text="Back", command=self.show_main_page).grid(row=5, columnspan=2, pady=5)

        self.chart_canvas = None
        self.pie_chart_display = tk.Frame(self.pie_chart_frame)
        self.pie_chart_display.pack(expand=True, fill=tk.BOTH)

    def add_result_column(self):
        var = tk.StringVar()
        self.result_col_vars.append(var)
        dropdown = ttk.Combobox(self.result_col_inner, textvariable=var, state="readonly", width=20)
        dropdown["values"] = list(self.df.columns) if self.df is not None else []
        dropdown.pack(pady=2, anchor="w")
        self.result_col_dropdowns.append(dropdown)

    def remove_result_column(self):
        if self.result_col_dropdowns:
            dropdown = self.result_col_dropdowns.pop()
            dropdown.destroy()
            self.result_col_vars.pop()

    def create_salary_page(self):
        tk.Label(self.salary_frame, text="Salary Calculator", font=("Arial", 14, "bold")).pack(pady=10)

        setting_frame = tk.Frame(self.salary_frame)
        setting_frame.pack(pady=5)

        tk.Label(setting_frame, text="Salary Type").grid(row=0, column=0)
        tk.Label(setting_frame, text="Work Hours Column").grid(row=1, column=0)
        tk.Label(setting_frame, text="Rate Column").grid(row=2, column=0)
        tk.Label(setting_frame, text="Currency").grid(row=3, column=0)
        tk.Label(setting_frame, text="Result Columns").grid(row=4, column=0)

        self.salary_type_var = tk.StringVar(value="hourly")
        self.work_hours_col_var = tk.StringVar()
        self.rate_col_var = tk.StringVar()
        self.currency_var = tk.StringVar(value="MYR")

        self.salary_type_dropdown = ttk.Combobox(setting_frame, textvariable=self.salary_type_var, state="readonly",
                                                 values=["hourly", "daily", "monthly"])
        self.salary_type_dropdown.grid(row=0, column=1)

        self.work_hours_dropdown = ttk.Combobox(setting_frame, textvariable=self.work_hours_col_var, state="readonly")
        self.work_hours_dropdown.grid(row=1, column=1)

        self.rate_col_dropdown = ttk.Combobox(setting_frame, textvariable=self.rate_col_var, state="readonly")
        self.rate_col_dropdown.grid(row=2, column=1)

        self.currency_dropdown = ttk.Combobox(setting_frame, textvariable=self.currency_var, state="readonly",
                                              values=["MYR", "IDR", "USD"])
        self.currency_dropdown.grid(row=3, column=1)

        # Scrollable result column area
        self.result_col_canvas = tk.Canvas(setting_frame, height=100)
        self.result_col_scrollbar = ttk.Scrollbar(setting_frame, orient="vertical",
                                                  command=self.result_col_canvas.yview)
        self.result_col_inner = tk.Frame(self.result_col_canvas)

        self.result_col_inner.bind("<Configure>", lambda e: self.result_col_canvas.configure(
            scrollregion=self.result_col_canvas.bbox("all")))
        self.result_col_canvas.create_window((0, 0), window=self.result_col_inner, anchor="nw")
        self.result_col_canvas.configure(yscrollcommand=self.result_col_scrollbar.set)

        self.result_col_canvas.grid(row=4, column=1, sticky="nsew")
        self.result_col_scrollbar.grid(row=4, column=2, sticky="ns")

        button_frame = tk.Frame(setting_frame)
        button_frame.grid(row=5, column=1, columnspan=2, sticky="w", pady=5)

        tk.Button(button_frame, text="+ Add Column", command=self.add_result_column).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="- Remove Column", command=self.remove_result_column).pack(side=tk.LEFT, padx=2)

        tk.Button(setting_frame, text="Calculate Salary", command=self.calculate_salary).grid(row=6, columnspan=2,
                                                                                              pady=10)
        tk.Button(setting_frame, text="Back", command=self.show_main_page).grid(row=7, columnspan=2, pady=5)

        self.salary_output_frame = tk.Frame(self.salary_frame)
        self.salary_output_frame.pack(expand=True, fill=tk.BOTH)

        self.salary_tree = ttk.Treeview(self.salary_output_frame, show="headings")
        self.salary_tree.pack(expand=True, fill=tk.BOTH)

        self.result_col_vars = []
        self.result_col_dropdowns = []
        self.add_result_column()  # Add initial dropdown

    def show_main_page(self):
        self.analysis_frame.pack_forget()
        self.pie_chart_frame.pack_forget()
        self.salary_frame.pack_forget()
        self.main_frame.pack(expand=True, fill=tk.BOTH)

    def show_analysis_page(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load an Excel file first.")
            return
        self.column_dropdown["values"] = list(self.df.columns)
        self.main_frame.pack_forget()
        self.pie_chart_frame.pack_forget()
        self.salary_frame.pack_forget()
        self.analysis_frame.pack(expand=True, fill=tk.BOTH)

    def show_pie_chart_page(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load an Excel file first.")
            return

        cols = list(self.df.columns)
        with_empty = [""] + cols

        self.col1_dropdown["values"] = with_empty
        self.col2_dropdown["values"] = with_empty

        self.col1_var.set("")
        self.col2_var.set("")

        self.main_frame.pack_forget()
        self.analysis_frame.pack_forget()
        self.salary_frame.pack_forget()
        self.pie_chart_frame.pack(expand=True, fill=tk.BOTH)

    def show_salary_page(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load an Excel file first.")
            return
        cols = list(self.df.columns)
        self.work_hours_dropdown["values"] = cols
        self.rate_col_dropdown["values"] = cols
        for combo in self.result_col_frame.winfo_children():
            if isinstance(combo, ttk.Combobox):
                combo["values"] = cols

        self.main_frame.pack_forget()
        self.analysis_frame.pack_forget()
        self.pie_chart_frame.pack_forget()
        self.salary_frame.pack(expand=True, fill=tk.BOTH)

        # Update dropdown values for each result column
        for dropdown in self.result_col_dropdowns:
            dropdown["values"] = cols


    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.excel_file_path = file_path
            self.update_file_label()
            self.load_excel()

    def update_file_label(self):
        if self.excel_file_path:
            file_name = os.path.basename(self.excel_file_path)
            self.file_label.config(text=f"Selected File: {file_name}", fg="blue")

    def load_excel(self):
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            self.file_label.config(text="Error: No valid file selected", fg="red")
            return

        self.df = pd.read_excel(self.excel_file_path, sheet_name=0)
        self.update_table(self.df)

    def update_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=150)
        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

    def update_analysis_table(self, df):
        self.analysis_tree.delete(*self.analysis_tree.get_children())
        self.analysis_tree["columns"] = list(df.columns)
        for col in df.columns:
            self.analysis_tree.heading(col, text=col)
            self.analysis_tree.column(col, anchor="center", width=150)
        for _, row in df.iterrows():
            self.analysis_tree.insert("", tk.END, values=list(row))

    def filter_data(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return

        column = self.column_var.get()
        value = self.filter_entry.get().strip()

        if not column:
            messagebox.showerror("Error", "Please select a column.")
            return

        if value.lower() == "all":
            filtered_df = self.df.copy()
        else:
            filtered_df = self.df[self.df[column].astype(str).str.contains(value, case=False, na=False)]

        self.filtered_df = filtered_df

        if filtered_df.empty:
            messagebox.showinfo("No Results", "No matching data found.")
        else:
            self.update_analysis_table(filtered_df)

    def calculate_total(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return

        column = self.column_var.get()
        if not column:
            messagebox.showerror("Error", "Please select a column.")
            return

        try:
            total = self.df[column].sum()
            messagebox.showinfo("Total", f"Total sum of '{column}': {total}")
        except Exception:
            messagebox.showerror("Error", "Selected column is not numeric.")

    def generate_pie_chart(self):
        col1 = self.col1_var.get()
        col2 = self.col2_var.get()

        try:
            df = self.df.copy()

            if col1 and col2:
                df["combined"] = df[col1].astype(str) + " - " + df[col2].astype(str)
            elif col1:
                df["combined"] = df[col1].astype(str)
            elif col2:
                df["combined"] = df[col2].astype(str)
            else:
                df["combined"] = "All Data"

            pie_data = df["combined"].value_counts()

            fig, ax = plt.subplots(figsize=(6, 6))
            ax.pie(pie_data, labels=pie_data.index, autopct="%1.1f%%", startangle=140, textprops={"fontsize": 8})
            ax.set_title("Pie Chart")

            if self.chart_canvas:
                self.chart_canvas.get_tk_widget().destroy()

            self.chart_canvas = FigureCanvasTkAgg(fig, master=self.pie_chart_display)
            self.chart_canvas.draw()
            self.chart_canvas.get_tk_widget().pack()

            fig.savefig("pie_chart_output.jpg")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate chart: {e}")

    def calculate_salary(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load data first.")
            return

        try:
            work_col = self.work_hours_col_var.get()
            rate_col = self.rate_col_var.get()
            salary_type = self.salary_type_var.get()
            currency = self.currency_var.get()

            df = self.df.copy()

            # Ensure that the work hours and rate columns are numeric
            df[work_col] = pd.to_numeric(df[work_col], errors='coerce', downcast='float')
            df[rate_col] = pd.to_numeric(df[rate_col], errors='coerce', downcast='float')

            # Validate that work hours and rate columns are numeric
            if df[work_col].isnull().any() or df[rate_col].isnull().any():
                messagebox.showerror("Error", "Work hours and rate columns must contain numeric values.")
                return

            # Reset the salary treeview
            self.salary_tree["columns"] = []
            self.salary_tree.delete(*self.salary_tree.get_children())

            # Apply salary calculation based on type
            if salary_type == "daily":
                df["Salary"] = df[work_col] * df[rate_col]
            elif salary_type == "monthly":
                df["Salary"] = df[rate_col]
            else:  # hourly
                df["Salary"] = df[work_col] * df[rate_col]

            # Add the salary label with currency
            salary_label = f"Salary ({currency})"
            df[salary_label] = df["Salary"]

            # Format the salary
            df[salary_label] = df[salary_label].apply(lambda x: f"{currency} {x:,.2f}")

            selected_cols = [var.get() for var in self.result_col_vars if var.get() in df.columns]
            selected_cols += [salary_label]

            display_df = df[selected_cols]

            # Update the salary treeview
            self.salary_tree["columns"] = list(display_df.columns)
            for col in display_df.columns:
                self.salary_tree.heading(col, text=col)
                self.salary_tree.column(col, anchor="center", width=150)
            for _, row in display_df.iterrows():
                self.salary_tree.insert("", tk.END, values=list(row))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate salary: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
