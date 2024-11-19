import pandas as pd
import zipfile
import os


excel_file_path = 'X:/AUR/23.10.2024/poze existente.xlsx'
zip_directory = 'X:/AUR/23.10.2024'

# Read the Excel file
df = pd.read_excel(excel_file_path)

# Print the column names to check their actual names
print("Column names in the Excel file:", df.columns.tolist())

# Identify items to delete based on the criteria
# Replace 'Your_Column_B_Name', 'Your_Column_C_Name', and 'Your_Column_A_Name' with actual column names
items_to_delete = set(df.loc[(df['B'] == 'No') | (df['C'].isnull()), 'A'])

# Loop through each zip file in the specified directory
for zip_filename in os.listdir(zip_directory):
    if zip_filename.endswith('.zip'):
        zip_file_path = os.path.join(zip_directory, zip_filename)
        
        with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
            # Read the contents of the zip file into a dictionary
            zip_contents = {file: zip_file.read(file) for file in zip_file.namelist()}
        
        # Create a list of files to delete
        files_to_delete = [file for file in zip_contents if file.endswith('.jpg') and file[:-4] in items_to_delete]
        
        # Define the modified zip file path
        modified_zip_file_path = os.path.join(zip_directory, zip_filename.replace('.zip', '_modified.zip'))
        
        # Create a new ZIP file excluding the unwanted files
        with zipfile.ZipFile(modified_zip_file_path, 'w') as new_zip:
            for file, content in zip_contents.items():
                if file not in files_to_delete:
                    new_zip.writestr(file, content)

        # Remove the old zip file
        if os.path.exists(zip_file_path):
            os.remove(zip_file_path)
        
        # Rename the modified ZIP file to the original name
        if os.path.exists(modified_zip_file_path):
            os.rename(modified_zip_file_path, zip_file_path)

print("Deletion of .jpg files based on Excel criteria completed.")
