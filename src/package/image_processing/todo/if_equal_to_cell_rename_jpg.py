# cauta intr-un folder cu foldere cu imagini .jpg
# daca gaseste imagine cu numele care coincide cu coloana B din excel,
# redenumeste imaginea cu valoarea din coloana A de pe acelasi rand din excel


import os
import pandas as pd

date_folder = "25.10.2024"
folder_path = f'X:\\AUR\\{date_folder}\\image_problem\\imagini'  # Path to your main folder with subfolders
excel_path = f'X:\\AUR\\{date_folder}\\image_problem\\items_system_db.xlsx'  # Path to your Excel file

# Load the Excel data
df = pd.read_excel(excel_path, usecols=[0, 1])  # Load columns by index (0 for A, 1 for B)
df = df.dropna()  # Drop rows with empty cells in either column

# Create a dictionary for easy lookup of B (index 1) -> A (index 0) values
name_map = dict(zip(df.iloc[:, 1], df.iloc[:, 0]))

# Traverse the folder and subfolders to find jpg files
for root, dirs, files in os.walk(folder_path):
    print("first for")
    for file in files:
        print("second for")
        if file.lower().endswith('.jpg'):
            print("first if")
            # Strip extension for name comparison
            file_name = os.path.splitext(file)[0]
            
            # Check if the filename is in the dictionary
            if file_name in name_map:
                print("second if")
                new_name = name_map[file_name]
                old_path = os.path.join(root, file)
                new_path = os.path.join(root, f"{new_name}.jpg")

                # Rename the file
                try:
                    os.rename(old_path, new_path)
                    print(f"Renamed '{file}' to '{new_name}.jpg'")
                except Exception as e:
                    print(f"Error renaming '{file}': {e}")
