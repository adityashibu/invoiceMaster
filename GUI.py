import os

import tkinter as tk
from tkinter import filedialog

from spire.xls import Workbook
from spire.common import FileFormat

# Global variable to keep track of the invoice number
invoice_number = 1

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
        global save_folder
        save_folder = foldername

def generate_invoices():
    global invoice_number
    template_file = "./template.xlsx"
    data_file = selected_data_label["text"].split(": ")[1]
    
    # Create folders if they don't exist
    sheets_folder = os.path.join(save_folder, "sheets")
    pdf_folder = os.path.join(save_folder, "pdf")
    os.makedirs(sheets_folder, exist_ok=True)
    os.makedirs(pdf_folder, exist_ok=True)

    # Load template
    workbook = Workbook()
    workbook.load_from_file(template_file)
    worksheet = workbook.active_sheet
    
    # Load data
    data_workbook = Workbook()
    data_workbook.load_from_file(data_file)
    data_worksheet = data_workbook.active_sheet

    for row_index, row in enumerate(data_worksheet.rows, start=1):
        # Update cell D9 with text from column B
        worksheet.range["D9"].text = row[1].value

        # Update cell E2 with text from column A
        worksheet.range["E2"].text = row[0].value
        worksheet.range["E3"].text = row[0].value
        
        # Update cell B16 with text from column C
        worksheet.range["B16"].text = row[2].value

        # Update cell C16 with text from column D
        worksheet.range["E16"].text = row[3].value

        # Update invoice number at E1
        worksheet.range["E1"].text = str(invoice_number)

        # Increment invoice number
        invoice_number += 1

        # Save the modified template to the 'sheets' folder
        invoice_filename = f"Invoice_{invoice_number}.xlsx"
        workbook.save_to_file(os.path.join(sheets_folder, invoice_filename))

        # Convert Excel to PDF
        pdf_filename = f"Invoice_{invoice_number}.pdf"
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        workbook.save_to_pdf(pdf_path, FileFormat.PDF)

    print("Invoices generated successfully!")

# Create main application window
root = tk.Tk()
root.title("Invoicing System")
root.geometry("700x250")  # Set the initial size of the window

# Create label to display selected data file
selected_data_label = tk.Label(root, text="")
selected_data_label.pack(pady=5)

# Create button for selecting data
data_button = tk.Button(root, text="Select Data Excel File", command=select_data, width=20)
data_button.pack(pady=5)

# Create label to display selected folder to save invoices
selected_save_folder_label = tk.Label(root, text="")
selected_save_folder_label.pack(pady=5)

# Create button for selecting folder to save invoices
save_button = tk.Button(root, text="Save to", command=select_save_folder, width=20)
save_button.pack(pady=5)

# Create button for generating invoices
generate_button = tk.Button(root, text="Generate Invoices", command=generate_invoices, width=20)
generate_button.pack(pady=5)

# Run the application
root.mainloop()