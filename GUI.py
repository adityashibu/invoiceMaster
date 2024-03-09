import tkinter as tk
from tkinter import filedialog

def select_data():
    filename = filedialog.askopenfilename(title="Select Data Excel File", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if filename:
        # Here, you can further process the selected Excel file
        print("Selected Data File:", filename)
        selected_data_label.config(text="Selected Data File: " + filename)

def select_save_folder():
    foldername = filedialog.askdirectory(title="Select Folder to Save Invoices")
    if foldername:
        # Here, you can further process the selected folder
        print("Selected Folder to Save Invoices:", foldername)
        selected_save_folder_label.config(text="Selected Folder to Save Invoices: " + foldername)

def generate_invoices():
    # Placeholder function for generating invoices
    print("Generating invoices...")

# Create main application window
root = tk.Tk()
root.title("Invoicing System")
root.geometry("500x300")  # Set the initial size of the window

# Create label to display selected data file
selected_data_label = tk.Label(root, text="")
selected_data_label.pack(pady=10)

# Create button for selecting data
data_button = tk.Button(root, text="Select Data Excel File", command=select_data, width=20)
data_button.pack(pady=10)

# Create label to display selected folder to save invoices
selected_save_folder_label = tk.Label(root, text="")
selected_save_folder_label.pack(pady=5)

# Create button for selecting folder to save invoices
save_button = tk.Button(root, text="Save to", command=select_save_folder, width=20)
save_button.pack(pady=5)

# Create button for generating invoices
generate_button = tk.Button(root, text="Generate Invoices", command=generate_invoices, width=20)
generate_button.pack(pady=10)

# Run the application
root.mainloop()
