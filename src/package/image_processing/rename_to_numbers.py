import os
import pandas as pd

# works only with unzipped folder
sub_folder = "111"
date_folder = "28.10.2024 - Copy"

excel_file = f'X:\\AUR\\{date_folder}\\image_problem\\num-copy_unzipped\\{sub_folder}\\041024.xlsx'
image_folder = f'X:\\AUR\\{date_folder}\\image_problem\\num-copy_unzipped\\{sub_folder}'
df = pd.read_excel(excel_file)

deleted_files = []

# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    correct_name = str(row[0])  # Get the correct name (column A)
    current_name = row[1] if pd.notna(row[1]) else ""  # Get the current name (column B) or set to empty string

    # Check if current_name is a valid string
    if isinstance(current_name, str) and current_name:
        _, ext = os.path.splitext(current_name)
    else:
        ext = '.jpg'  # Default extension if current_name is missing
        # new_image_path = os.path.join(image_folder, "!missing_name" + ext)
        # print(f"Missing current name for row {index+2}, assigning default name: {new_image_path}")
        print(f"Missing current name for row {index+2}, correct name {correct_name}")

    if ext.lower() != '.jpg':
        print(f"Skipping: {current_name} (not a .jpg file)")
        continue

    # Construct the full paths
    current_image_path = os.path.join(image_folder, current_name) if current_name else ""
    new_image_path = os.path.join(image_folder, correct_name + ext)

    # Skip renaming if the current name is already the correct name
    if current_image_path == new_image_path:
        # print(f"Skipping: {current_name} is already named correctly.")
        continue

    # Check if the current image exists
    if current_image_path and os.path.exists(current_image_path):
        # If the new file name already exists, delete the existing file
        if os.path.exists(new_image_path):
            os.remove(new_image_path)
            deleted_files.append(new_image_path)
            print(f"Deleted existing file: {new_image_path}")

        # Rename the file to the new name
        os.rename(current_image_path, new_image_path)
        print(f"Renamed: {current_name} -> {correct_name + ext}")
    else:
        print(f"File not found: {current_name}")

# Print all deleted files at the end
if deleted_files:
    print("\nDeleted files:")
    for file in deleted_files:
        print(file)
else:
    print("\nNo files were deleted.")

print("Renaming complete.")
