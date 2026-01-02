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
        self.root.geometry("1000x600")

        self.excel_file_path = None
        self.df = None
        self.filtered_df = None
        self.chart_canvas = None

        self.main_frame = tk.Frame(root)
        self.analysis_frame = tk.Frame(root)
        self.settings_frame = tk.Frame(root)
        self.pie_frame = tk.Frame(root)

        self.create_main_page()
        self.create_analysis_page()
        self.create_settings_page()
        self.create_pie_page()

        self.show_main_page()

    def create_main_page(self):
        frame = tk.Frame(self.main_frame)
        frame.pack(pady=10)

        self.load_button = tk.Button(frame, text="Refresh Data", command=self.load_excel)
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.browse_button = tk.Button(frame, text="Browse File", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        self.analyze_button = tk.Button(frame, text="Analyze", command=self.show_analysis_page)
        self.analyze_button.pack(side=tk.LEFT, padx=5)

        self.settings_button = tk.Button(frame, text="Settings", command=self.show_settings_page)
        self.settings_button.pack(side=tk.LEFT, padx=5)

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

        self.filter_button = tk.Button(control_frame, text="Filter", command=self.filter_data)
        self.filter_button.pack(side=tk.LEFT, padx=5)

        self.total_button = tk.Button(control_frame, text="Calculate Total", command=self.calculate_total)
        self.total_button.pack(side=tk.LEFT, padx=5)

        self.pie_button = tk.Button(control_frame, text="Show Pie Chart", command=self.show_pie_page)
        self.pie_button.pack(side=tk.LEFT, padx=5)

        self.back_button = tk.Button(control_frame, text="Back to Main", command=self.show_main_page)
        self.back_button.pack(side=tk.LEFT, padx=5)

        self.filter_note = tk.Label(self.analysis_frame, text="Type value or 'ALL' on the form", fg="gray")
        self.filter_note.pack(pady=(0, 10))

        self.analysis_table_frame = tk.Frame(self.analysis_frame)
        self.analysis_table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.analysis_tree = ttk.Treeview(self.analysis_table_frame, show="headings")
        self.analysis_tree.pack(expand=True, fill=tk.BOTH)

    def create_settings_page(self):
        tk.Label(self.settings_frame, text="Settings", font=("Arial", 14, "bold")).pack(pady=10)
        self.back_button = tk.Button(self.settings_frame, text="Back to Main", command=self.show_main_page)
        self.back_button.pack(pady=5)

    def create_pie_page(self):
        tk.Label(self.pie_frame, text="Pie Chart View", font=("Arial", 14, "bold")).pack(pady=10)

        self.pie_canvas_frame = tk.Frame(self.pie_frame)
        self.pie_canvas_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        self.pie_back_button = tk.Button(self.pie_frame, text="Back to Analysis", command=self.show_analysis_page)
        self.pie_back_button.pack(pady=10)

    def show_main_page(self):
        self.hide_all_frames()
        self.main_frame.pack(expand=True, fill=tk.BOTH)

    def show_analysis_page(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded. Please select an Excel file first.")
            return

        self.column_dropdown["values"] = list(self.df.columns)
        self.hide_all_frames()
        self.analysis_frame.pack(expand=True, fill=tk.BOTH)

    def show_settings_page(self):
        self.hide_all_frames()
        self.settings_frame.pack(expand=True, fill=tk.BOTH)

    def show_pie_page(self):
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showinfo("No Data", "Please apply a filter or select 'ALL' first.")
            return

        column = self.column_var.get()
        if not column:
            messagebox.showerror("Error", "Please select a column first.")
            return

        pie_data = self.filtered_df[column].value_counts()
        if pie_data.empty:
            messagebox.showinfo("No Data", "No data available for chart.")
            return

        fig, ax = plt.subplots()
        ax.pie(pie_data.values, labels=pie_data.index, autopct='%1.1f%%', startangle=90)
        ax.set_title(f"Pie Chart of {column}")

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()

        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.pie_canvas_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)

        self.hide_all_frames()
        self.pie_frame.pack(expand=True, fill=tk.BOTH)

    def hide_all_frames(self):
        self.main_frame.pack_forget()
        self.analysis_frame.pack_forget()
        self.settings_frame.pack_forget()
        self.pie_frame.pack_forget()

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
            if not value:
                messagebox.showerror("Error", "Please enter a value or 'ALL'.")
                return
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
            total = pd.to_numeric(self.df[column], errors='coerce').sum()
            messagebox.showinfo("Total", f"Total sum of '{column}': {total}")
        except Exception:
            messagebox.showerror("Error", "Selected column is not numeric.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
