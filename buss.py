import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb  # For modern themes

# File path to store bus data
FILE_PATH = 'buses.xlsx'

def load_data():
    try:
        return pd.read_excel(FILE_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Bus ID', 'Route', 'Driver', 'Insurance Expiry', 'Maintenance Due', 'Status'])
        df.to_excel(FILE_PATH, index=False)
        return df

def save_data(df):
    try:
        df.to_excel(FILE_PATH, index=False)
    except PermissionError:
        messagebox.showwarning("Permission Error", "Please close the Excel file and try again.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving data: {e}")

# Initialize DataFrame
df = load_data()

# Functions for GUI
def view_all_buses():
    for row in tree.get_children():
        tree.delete(row)
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

def add_or_update_bus():
    global df
    bus_id = entry_bus_id.get()
    route = entry_route.get()
    driver = entry_driver.get()
    insurance_expiry = entry_insurance.get()
    maintenance_due = entry_maintenance.get()
    status = combo_status.get()
    
    if not all([bus_id, route, driver, insurance_expiry, maintenance_due, status]):
        messagebox.showwarning("Input Error", "All fields must be filled!")
        return

    if bus_id in df['Bus ID'].values:
        df.loc[df['Bus ID'] == bus_id, ['Route', 'Driver', 'Insurance Expiry', 'Maintenance Due', 'Status']] = \
            [route, driver, insurance_expiry, maintenance_due, status]
    else:
        new_row = {
            'Bus ID': bus_id, 'Route': route, 'Driver': driver,
            'Insurance Expiry': insurance_expiry, 'Maintenance Due': maintenance_due, 'Status': status
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    save_data(df)
    view_all_buses()
    messagebox.showinfo("Success", "Bus details have been updated!")

def check_expired_insurance():
    expired = df[pd.to_datetime(df['Insurance Expiry'], errors='coerce') < datetime.now()]
    if expired.empty:
        messagebox.showinfo("Insurance Check", "No buses with expired insurance.")
    else:
        show_popup("Buses with Expired Insurance", expired)

def check_maintenance_due():
    due = df[pd.to_datetime(df['Maintenance Due'], errors='coerce') <= datetime.now()]
    if due.empty:
        messagebox.showinfo("Maintenance Check", "No buses needing maintenance.")
    else:
        show_popup("Buses Needing Maintenance", due)

def show_popup(title, data):
    popup = tb.Toplevel()
    popup.title(title)
    popup.geometry("800x400")
    tree_popup = ttk.Treeview(popup, columns=list(data.columns), show='headings', bootstyle="info")
    for col in data.columns:
        tree_popup.heading(col, text=col)
        tree_popup.column(col, width=150)
    tree_popup.pack(fill='both', expand=True)
    for _, row in data.iterrows():
        tree_popup.insert("", "end", values=list(row))

# GUI setup
app = tb.Window(themename="darkly")
app.title("Bus Management System")
app.geometry("900x700")

# Header
header = ttk.Label(app, text="Bus Management System", font=("Helvetica", 20), bootstyle="success")
header.pack(pady=20)

# Input Frame
input_frame = ttk.Frame(app, padding=10)
input_frame.pack(pady=10, fill="x")

ttk.Label(input_frame, text="Bus ID").grid(row=0, column=0, padx=5, pady=5)
entry_bus_id = ttk.Entry(input_frame, width=20)
entry_bus_id.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Route").grid(row=1, column=0, padx=5, pady=5)
entry_route = ttk.Entry(input_frame, width=20)
entry_route.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Driver").grid(row=2, column=0, padx=5, pady=5)
entry_driver = ttk.Entry(input_frame, width=20)
entry_driver.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Insurance Expiry (YYYY-MM-DD)").grid(row=3, column=0, padx=5, pady=5)
entry_insurance = ttk.Entry(input_frame, width=20)
entry_insurance.grid(row=3, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Maintenance Due (YYYY-MM-DD)").grid(row=4, column=0, padx=5, pady=5)
entry_maintenance = ttk.Entry(input_frame, width=20)
entry_maintenance.grid(row=4, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Status").grid(row=5, column=0, padx=5, pady=5)
combo_status = ttk.Combobox(input_frame, values=["Active", "Under Maintenance"], width=18)
combo_status.grid(row=5, column=1, padx=5, pady=5)

ttk.Button(input_frame, text="Add/Update Bus", bootstyle="primary", command=add_or_update_bus).grid(row=6, columnspan=2, pady=10)

# Table Frame
tree_frame = ttk.Frame(app)
tree_frame.pack(fill="both", expand=True, pady=10)

columns = ['Bus ID', 'Route', 'Driver', 'Insurance Expiry', 'Maintenance Due', 'Status']
tree = ttk.Treeview(tree_frame, columns=columns, show='headings', bootstyle="info")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.pack(fill="both", expand=True)

# Button Frame
button_frame = ttk.Frame(app, padding=10)
button_frame.pack(pady=10)

ttk.Button(button_frame, text="View All Buses", bootstyle="success", command=view_all_buses).grid(row=0, column=0, padx=10)
ttk.Button(button_frame, text="Check Expired Insurance", bootstyle="warning", command=check_expired_insurance).grid(row=0, column=1, padx=10)
ttk.Button(button_frame, text="Check Maintenance Due", bootstyle="danger", command=check_maintenance_due).grid(row=0, column=2, padx=10)
ttk.Button(button_frame, text="Exit", bootstyle="secondary", command=app.quit).grid(row=0, column=3, padx=10)

# Load initial data
view_all_buses()

app.mainloop()
