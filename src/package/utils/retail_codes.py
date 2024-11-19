import pandas as pd

# Define file path
file_path = 'X:/AUR/11.2024/06.11.2024/CATEGORIES WITH CM - Copy.xlsx'  # Replace with the path to your Excel file

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path, header=None)  # Load without headers to use indexes

# Iterate through rows and check for matches
for index, row in df.iterrows():
    # Column F (index 5): split into words to search
    words = str(row[5]).split()  # Split words in Column F
    for word in words:
        if word.lower() in str(row[11]).lower():  # Column L (index 11) - case-insensitive search
            df.at[index, 10] = row[11]  # Store match in Column K (index 10)
            break  # Exit loop once a match is found

df.to_excel('X:/AUR/11.2024/06.11.2024/updated_file.xlsx', index=False, header=False)  # Save as a new file

print("Processing complete. Check 'updated_file.xlsx' for results.")
