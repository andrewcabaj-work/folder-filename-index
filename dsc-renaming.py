import os
import pandas as pd

def rename_files_in_subfolders(base_directory):
    renamed_filenames = []

    # Get the parent directory of the base directory
    parent_directory = os.path.dirname(base_directory)

    # Iterate through each subfolder in the base directory
    for subfolder_name in os.listdir(base_directory):
        subfolder_path = os.path.join(base_directory, subfolder_name)
        
        if os.path.isdir(subfolder_path):
            # Iterate through each file in the subfolder
            for filename in os.listdir(subfolder_path):
                if "Thumbs" in filename:
                    continue  # Skip files containing "Thumbs"
                
                file_path = os.path.join(subfolder_path, filename)
                
                if os.path.isfile(file_path):
                    # Split the filename and extension
                    name, ext = os.path.splitext(filename)
                    
                    # Construct the new filename with the extension
                    new_filename = f"[Doc Production] {subfolder_name} ({name}){ext}"
                    new_file_path = os.path.join(parent_directory, new_filename)
                    
                    # Move the file to the parent directory
                    os.rename(file_path, new_file_path)
                    print(f"Moved and renamed '{file_path}' to '{new_file_path}'")
                    
                    # Store the new filename without the extension
                    renamed_filenames.append(os.path.splitext(new_filename)[0])

    # Export the renamed filenames without extension to an Excel spreadsheet
    df = pd.DataFrame(renamed_filenames, columns=["Filename"])
    export_path = os.path.join(parent_directory, "dsc-filename-index.xlsx")
    df.to_excel(export_path, index=False)
    print(f"Exported renamed filenames to '{export_path}'")

# Define the base directory
base_directory = r"enter path here"

# Call the function to rename files and export filenames
rename_files_in_subfolders(base_directory)
