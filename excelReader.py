import os
import pandas as pd
import tkinter as tk
from tkinter import ttk

# Set the exact file path
EXCEL_FILE_PATH = r"E:\SEPUH DATA\contoh data analisis python.xlsx"

class ExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Viewer")
        self.root.geometry("800x500")

        # Frame for buttons
        frame = tk.Frame(root)
        frame.pack(pady=10)

        self.load_button = tk.Button(frame, text="Refresh Data", command=self.load_excel)
        self.load_button.pack(side=tk.LEFT, padx=5)

        # Treeview (Table) for displaying Excel data
        self.tree = ttk.Treeview(root, show="headings")
        self.tree.pack(expand=True, fill=tk.BOTH)

    def load_excel(self):
        """Loads the specific Excel file and displays it."""
        if not os.path.exists(EXCEL_FILE_PATH):
            print(f"File not found: {EXCEL_FILE_PATH}")
            return

        self.display_excel(EXCEL_FILE_PATH)

    def display_excel(self, file_path):
        """Reads the first sheet of the Excel file and displays it in the GUI."""
        df = pd.read_excel(file_path, sheet_name=0)  # Read first sheet
        self.update_table(df)

    def update_table(self, df):
        """Updates the Treeview table with DataFrame data."""
        # Clear existing data
        self.tree.delete(*self.tree.get_children())

        # Define columns
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        # Insert rows
        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

        self.tree.pack(expand=True, fill=tk.BOTH)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewer(root)
    app.load_excel()  # Automatically load on startup
    root.mainloop()
