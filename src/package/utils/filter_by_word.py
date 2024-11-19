import pandas as pd

# if finds the word "BROS " the row remains, else it deletes the row

# Load your Excel file
file_path = "X:/AUR/29.10.2024/Items by Location Matrix 16-09-2024.xlsx"
df = pd.read_excel(file_path)

# Define the column you want to filter
column_to_filter = 'Description'

# Filter rows that contain "BROS " in the specified column
df = df[df[column_to_filter].str.contains("BROS ", na=False)]

# Display or save the filtered data
print(df)
# Save the filtered DataFrame back to the original file (or to a new file)
df.to_excel(file_path, index=False)
