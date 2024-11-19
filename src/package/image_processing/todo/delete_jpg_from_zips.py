import pandas as pd
import zipfile
import os

# this deletes .jpg files, going through zip folders based on column B and C,
# if B is No or C is empty, then it deletes the file. Ex:
'''
ITM009969	Yes	68767
ITM009972	Yes	
ITM010103	No
ITM011116	No	83900
'''

# Yes = am poza
# No = nu are poza


excel_file_path = 'X:/AUR/23.10.2024/poze existente.xlsx'
zip_directory = 'X:/AUR/23.10.2024'

df = pd.read_excel(excel_file_path)

# Identify items to delete based on the criteria
items_to_delete = set(df.loc[(df['B'] == 'No') | (df['C'].isnull()), 'A'])

# Loop through each zip file in the specified directory
for zip_filename in os.listdir(zip_directory):
    if zip_filename.endswith('.zip'):
        zip_file_path = os.path.join(zip_directory, zip_filename)
        
        with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
            # Create a list to hold the names of files to be deleted
            files_to_delete = [file for file in zip_file.namelist() if file.endswith('.jpg') and file[:-4] in items_to_delete]
        
        # Create a new ZIP file excluding the unwanted files
        with zipfile.ZipFile(zip_file_path.replace('.zip', '_modified.zip'), 'w') as new_zip:
            for file in zip_file.namelist():
                if file not in files_to_delete:
                    new_zip.writestr(file, zip_file.read(file))
        
        # Optionally, remove the old zip file and rename the modified one
        os.remove(zip_file_path)
        os.rename(zip_file_path.replace('.zip', '_modified.zip'), zip_file_path)

print("Deletion of .jpg files based on Excel criteria completed.")
