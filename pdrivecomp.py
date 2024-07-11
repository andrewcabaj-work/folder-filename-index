import os
from openpyxl import Workbook
import ctypes

def is_hidden(filepath):
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(filepath))
        assert attrs != -1
        result = bool(attrs & 2)
    except (AttributeError, AssertionError):
        result = False
    return result

# Ask the user for the folder path
folder_path = input("Please enter the folder path: ")

# List all files in the folder
filenames = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

# Create a new Excel workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Write filenames to the Excel sheet
for idx, filename in enumerate(filenames, start=1):
    name_without_extension = os.path.splitext(filename)[0]
    ws[f"A{idx}"] = name_without_extension    
    
# Save the workbook
downloads_folder = "C:\\Users\\andrew.cabaj\\Downloads"
folder_name = os.path.basename(folder_path)  # Get the last part of the folder_path as folder name
safe_folder_name = folder_name.replace(':', '').replace('\\', '').replace('/', '')  # Remove potentially illegal characters
excel_filename = f"Filenames - {safe_folder_name}.xlsx"
excel_path = os.path.join(downloads_folder, excel_filename)  # Full path to save the file

wb.save(excel_path)

print(f"Excel spreadsheet '{excel_path}' has been created with the filenames.")
