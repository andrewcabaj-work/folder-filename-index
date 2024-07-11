import os
from openpyxl import Workbook
import ctypes
import tkinter as tk
from tkinter import filedialog

def is_hidden(filepath):
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(filepath))
        assert attrs != -1
        result = bool(attrs & 2)
    except (AttributeError, AssertionError):
        result = False
    return result

def generate_excel(folder_path):
    filenames = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f)) and not is_hidden(os.path.join(folder_path, f))]
    wb = Workbook()
    ws = wb.active
    for idx, filename in enumerate(filenames, start=1):
        name_without_extension = os.path.splitext(filename)[0]
        ws[f"A{idx}"] = name_without_extension
    downloads_folder = "C:\\Users\\andrew.cabaj\\Downloads"
    folder_name = os.path.basename(folder_path)
    safe_folder_name = folder_name.replace(':', '').replace('\\', '').replace('/', '')
    excel_filename = f"Filenames - {safe_folder_name}.xlsx"
    excel_path = os.path.join(downloads_folder, excel_filename)
    wb.save(excel_path)
    print(f"Excel spreadsheet '{excel_path}' has been created with the filenames.")

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_path)

def on_generate_button_clicked():
    folder_path = folder_path_entry.get()
    if folder_path:
        generate_excel(folder_path)
    else:
        print("Please select a folder.")

# Set up the Tkinter window
root = tk.Tk()
root.title("P Drive - Filename Compiler")

# Create and pack the folder path entry widget
folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.grid(row=0, column=0, columnspan=2, padx=10, pady=0)

# Create and pack the browse button
browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.grid(row=0, column=2, padx=(0,5), pady=5)

# Create and pack the generate button
generate_button = tk.Button(root, text="Generate Spreadsheet", command=on_generate_button_clicked)
generate_button.grid(row=1, column=0, columnspan=3, padx=0, pady=(0,5))

root.mainloop()