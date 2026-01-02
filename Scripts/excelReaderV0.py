import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Viewer")
        self.root.geometry("900x500")

        # Default Excel File Path
        self.excel_file_path = None
        self.df = None  # Store dataframe

        # Create Main, Analysis, and Settings Pages
        self.main_frame = tk.Frame(root)
        self.analysis_frame = tk.Frame(root)
        self.settings_frame = tk.Frame(root)

        self.create_main_page()
        self.create_analysis_page()
        self.create_settings_page()

        # Show Main Page
        self.show_main_page()

    def create_main_page(self):
        """Create the Main Page layout."""
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

        # Create table frame with scrollbars
        table_frame = tk.Frame(self.main_frame)
        table_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        self.tree = ttk.Treeview(table_frame, show="headings")

        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree.pack(expand=True, fill=tk.BOTH)

    def create_analysis_page(self):
        """Create the Data Analysis Page layout."""
        tk.Label(self.analysis_frame, text="Data Analysis", font=("Arial", 14, "bold")).pack(pady=10)

        # Column selection dropdown
        self.column_var = tk.StringVar()
        self.column_dropdown = ttk.Combobox(self.analysis_frame, textvariable=self.column_var, state="readonly")
        self.column_dropdown.pack(pady=5)

        # Filter entry
        self.filter_entry = tk.Entry(self.analysis_frame)
        self.filter_entry.pack(pady=5)

        # Filter button
        self.filter_button = tk.Button(self.analysis_frame, text="Filter", command=self.filter_data)
        self.filter_button.pack(pady=5)

        # Calculate total button
        self.total_button = tk.Button(self.analysis_frame, text="Calculate Total", command=self.calculate_total)
        self.total_button.pack(pady=5)

        # Back button
        self.back_button = tk.Button(self.analysis_frame, text="Back to Main", command=self.show_main_page)
        self.back_button.pack(pady=10)

    def create_settings_page(self):
        """Create the Settings Page layout."""
        tk.Label(self.settings_frame, text="Settings", font=("Arial", 14, "bold")).pack(pady=10)
        self.back_button = tk.Button(self.settings_frame, text="Back to Main", command=self.show_main_page)
        self.back_button.pack(pady=5)

    def show_main_page(self):
        """Switch to the Main Page."""
        self.analysis_frame.pack_forget()
        self.settings_frame.pack_forget()
        self.main_frame.pack(expand=True, fill=tk.BOTH)

    def show_analysis_page(self):
        """Switch to the Data Analysis Page."""
        if self.df is None:
            messagebox.showerror("Error", "No data loaded. Please select an Excel file first.")
            return

        self.column_dropdown["values"] = list(self.df.columns)
        self.analysis_frame.pack(expand=True, fill=tk.BOTH)
        self.main_frame.pack_forget()

    def show_settings_page(self):
        """Switch to the Settings Page."""
        self.main_frame.pack_forget()
        self.settings_frame.pack(expand=True, fill=tk.BOTH)

    def browse_file(self):
        """Open a file dialog to select an Excel file."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.excel_file_path = file_path
            self.update_file_label()
            self.load_excel()

    def update_file_label(self):
        """Update the label to show the selected file name."""
        if self.excel_file_path:
            file_name = os.path.basename(self.excel_file_path)
            self.file_label.config(text=f"Selected File: {file_name}")

    def load_excel(self):
        """Loads the selected Excel file and displays it."""
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            self.file_label.config(text="Error: No valid file selected", fg="red")
            return

        self.file_label.config(fg="blue")
        self.df = pd.read_excel(self.excel_file_path, sheet_name=0)
        self.update_table(self.df)

    def update_table(self, df):
        """Updates the Treeview table with DataFrame data."""
        self.tree.delete(*self.tree.get_children())

        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=150)

        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

        self.tree.pack(expand=True, fill=tk.BOTH)

    def filter_data(self):
        """Filters data based on the selected column and entered value."""
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return

        column = self.column_var.get()
        value = self.filter_entry.get().strip()

        if not column or not value:
            messagebox.showerror("Error", "Please select a column and enter a value to filter.")
            return

        filtered_df = self.df[self.df[column].astype(str).str.contains(value, case=False, na=False)]
        if filtered_df.empty:
            messagebox.showinfo("No Results", "No matching data found.")
        else:
            self.update_table(filtered_df)

    def calculate_total(self):
        """Calculates the total of a numeric column."""
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return

        column = self.column_var.get()
        if not column:
            messagebox.showerror("Error", "Please select a column to calculate the total.")
            return

        try:
            total = self.df[column].sum()
            messagebox.showinfo("Total", f"Total sum of '{column}': {total}")
        except Exception:
            messagebox.showerror("Error", "Selected column is not numeric.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
