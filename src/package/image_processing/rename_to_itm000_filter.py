import os
import pandas as pd


sub_folder = "111"
date_folder = "28.10.2024 - Copy"

excel_file = f'X:/AUR/{date_folder}/image_problem/alex_report.xlsx'
image_folder = f'X:/AUR/{date_folder}/image_problem/num-copy_unzipped/{sub_folder}'
df = pd.read_excel(excel_file)

deleted_files = []

for index, row in df.iterrows():
    correct_name = str(row.iloc[0])  # column A (ITM number)
    current_name = str(row.iloc[4]) + '.jpg'  # column E (Attrib 1 Code) with .jpg extension added
    have_picture = str(row.iloc[2])  # column C (pictures)

    current_image_path = os.path.join(image_folder, current_name)
    new_image_path = os.path.join(image_folder, correct_name + '.jpg')

    # Skip renaming if the file already has the correct name
    if current_image_path == new_image_path:
        print(f"Skipping: {current_name} is already named correctly.")
        continue

    if os.path.exists(current_image_path):
        # Check if the have_picture column matches the specified value
        if have_picture != "{00000000-0000-0000-0000-000000000000}":
            os.remove(current_image_path)
            deleted_files.append(current_image_path)
            print(f"Deleted file due to no picture requirement: {current_image_path}")
            continue  # Skip to the next row since we don't need to rename

        # If the new file name already exists, delete the existing file
        if os.path.exists(new_image_path):
            os.remove(new_image_path)
            deleted_files.append(new_image_path)
            print(f"Deleted existing file: {new_image_path}")

        # Rename the current image to the correct name
        os.rename(current_image_path, new_image_path)
        print(f"Renamed: {current_name} -> {correct_name}.jpg")
    # else:
        # print(f"File not found: {current_name}")

# Summary of deleted files
if input("Show deleted files(Y/N): ") in ["Y", "y", "yes", "YES", "Yes", "YEs"]:
    if deleted_files:
        print("\nDeleted files:")
        for file in deleted_files:
            print(file)
    else:
        print("\nNo files were deleted.")
else:
    print("Total files deleted: ", len(deleted_files))

print("Renaming complete.")
