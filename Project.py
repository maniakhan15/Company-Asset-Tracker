import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import csv
import os
from datetime import datetime


DB_PATH = r"C:\Project\AssetDB1.accdb"


class AssetRegisterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Company Asset Register System")
        self.root.geometry("950x650")
        self.root.configure(bg="#e8eaf6") 
        self.setup_ui()
        
       
        self.conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=" + DB_PATH + ";"
        )
        self.check_db_connection()
        self.refresh_data()

    def check_db_connection(self):
        try:
            conn = pyodbc.connect(self.conn_str)
            conn.close()
            self.status_lbl.config(text="Database Status: Connected (MS Access)", fg="green")
        except Exception as e:
            self.status_lbl.config(text=f"Error: {str(e)}", fg="red")
            messagebox.showerror("Connection Error", 
                "Cannot connect to MS Access.\n\n1. Check the DB_PATH in line 14.\n2. Ensure the Access file is CLOSED.")

    def setup_ui(self):
        # Header
        header = tk.Frame(self.root, bg="#3f51b5", height=70)
        header.pack(fill="x")
        tk.Label(header, text="Fixed Asset Register", font=("Segoe UI", 24, "bold"), bg="#3f51b5", fg="white").pack(pady=10)

        # Main Layout
        main_frame = tk.Frame(self.root, bg="#e8eaf6")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # LEFT SIDE: Input Form
        input_frame = tk.LabelFrame(main_frame, text="New Asset Entry", bg="white", font=("Arial", 11, "bold"), padx=15, pady=15)
        input_frame.pack(side="left", fill="y", padx=(0, 20))

        # Asset Name
        tk.Label(input_frame, text="Asset Name / Model:", bg="white").pack(anchor="w", pady=(5, 0))
        self.name_entry = tk.Entry(input_frame, width=35)
        self.name_entry.pack(pady=(2, 10))

        # Serial Number
        tk.Label(input_frame, text="Serial Number:", bg="white").pack(anchor="w")
        self.serial_entry = tk.Entry(input_frame, width=35)
        self.serial_entry.pack(pady=(2, 10))

        # Category
        tk.Label(input_frame, text="Category:", bg="white").pack(anchor="w")
        self.cat_combo = ttk.Combobox(input_frame, values=["IT Equipment", "Furniture", "Machinery", "Vehicles", "Office Supplies"], width=32)
        self.cat_combo.pack(pady=(2, 10))

        # Cost
        tk.Label(input_frame, text="Cost ($):", bg="white").pack(anchor="w")
        self.cost_entry = tk.Entry(input_frame, width=35)
        self.cost_entry.pack(pady=(2, 10))

        # Assigned To
        tk.Label(input_frame, text="Assigned To (Employee/Dept):", bg="white").pack(anchor="w")
        self.assign_entry = tk.Entry(input_frame, width=35)
        self.assign_entry.pack(pady=(2, 20))

        # Buttons
        tk.Button(input_frame, text="Add Asset", command=self.add_asset, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), height=2, width=30).pack(pady=5)
        tk.Button(input_frame, text="Export to Excel", command=self.export_csv, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), height=2, width=30).pack(pady=5)

        # RIGHT SIDE: Treeview (Table)
        list_frame = tk.LabelFrame(main_frame, text="Asset Inventory", bg="white", font=("Arial", 11, "bold"))
        list_frame.pack(side="right", fill="both", expand=True)

        cols = ("ID", "Name", "Serial", "Category", "Cost", "Date", "Assigned")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings")
        
        # Define Headings
        for col in cols:
            self.tree.heading(col, text=col)
            
        # Column Widths
        self.tree.column("ID", width=30)
        self.tree.column("Name", width=120)
        self.tree.column("Serial", width=80)
        self.tree.column("Category", width=90)
        self.tree.column("Cost", width=60)
        self.tree.column("Date", width=80)
        self.tree.column("Assigned", width=100)

        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # Status Bar
        self.status_lbl = tk.Label(self.root, text="System Ready", bd=1, relief=tk.SUNKEN, anchor="w")
        self.status_lbl.pack(side="bottom", fill="x")

    def add_asset(self):
        name = self.name_entry.get()
        serial = self.serial_entry.get()
        cat = self.cat_combo.get()
        cost = self.cost_entry.get()
        assigned = self.assign_entry.get()
        date_str = datetime.now().strftime("%Y-%m-%d")

        if not name or not cost:
            messagebox.showwarning("Missing Info", "Please enter Asset Name and Cost.")
            return

        try:
            cost_val = float(cost)
        except ValueError:
            messagebox.showerror("Invalid Input", "Cost must be a number.")
            return

        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            
            sql = "INSERT INTO AssetLog (AssetName, SerialNumber, Category, PurchaseDate, Cost, AssignedTo) VALUES (?, ?, ?, ?, ?, ?)"
            cursor.execute(sql, (name, serial, cat, date_str, cost_val, assigned))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", "Asset Registered Successfully!")
            self.clear_inputs()
            self.refresh_data()
            
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not save: {e}")

    def refresh_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM AssetLog ORDER BY ID DESC")
            rows = cursor.fetchall()
            for row in rows:
                # Format the row for the treeview
                self.tree.insert("", "end", values=(row[0], row[1], row[2], row[3], f"${row[5]}", row[4], row[6]))
            conn.close()
        except Exception as e:
            pass

    def export_csv(self):
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM AssetLog")
            rows = cursor.fetchall()
            conn.close()

            filename = "Asset_Register_Export.csv"
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["ID", "Name", "Serial", "Category", "Date", "Cost", "Assigned To"])
                for row in rows:
                    writer.writerow([x for x in row])
            
            messagebox.showinfo("Export Success", f"File saved as {filename}\nOpening now...")
            os.system(f"start {filename}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def clear_inputs(self):
        self.name_entry.delete(0, tk.END)
        self.serial_entry.delete(0, tk.END)
        self.cost_entry.delete(0, tk.END)
        self.assign_entry.delete(0, tk.END)
        self.cat_combo.set('')

if __name__ == "__main__":
    root = tk.Tk()
    app = AssetRegisterApp(root)
    root.mainloop()