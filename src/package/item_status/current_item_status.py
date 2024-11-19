import pandas as pd
from datetime import datetime

# item status, citeste doua coloane, una cu ITM00100 si una cu data trecuta,
# sterge data care depaseste ziua de azi, sterge duplicatele la ITM
# pastrand ultima data actuala care nu depaseste data curenta 

file_path = 'X:/AUR/11.2024/05.11.2024/LSC Item Status Link.xlsx'  # Replace with your file path
output_path = 'X:/AUR/11.2024/05.11.2024/LSC Item Status Link.xlsx'  # Change output file name to avoid overwriting

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()
column_a = 'A'
column_f = 'F'
index_a = ord(column_a) - ord('A')
index_f = ord(column_f) - ord('A')

# Step 1: Convert column F to datetime
df.iloc[:, index_f] = pd.to_datetime(df.iloc[:, index_f], errors='coerce')
current_date = pd.to_datetime(datetime.now().date())

# Step 2: Delete rows where the date in column F is later than the current date
df = df[df.iloc[:, index_f] <= current_date]

# Step 3: Remove duplicates based on column A, keeping the row with the most recent date in column F
df = df.loc[df.groupby(df.columns[index_a])[df.columns[index_f]].idxmax()]

# Convert the date in column F to string format and slice to keep only the date
df.iloc[:, index_f] = df.iloc[:, index_f].astype(str).str[:10]  # Keep only the date part

# Save the cleaned DataFrame back to Excel
df.to_excel(output_path, index=False)

print("Cleaning complete. The cleaned file is saved as:", output_path)
