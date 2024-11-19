# pentru a obtine Excelul cu Retail Product Code si CM's:
# in BC > Configuration Packages > ALEX_RAPORT > Excel > Export to Excel

import pandas as pd

file1_path = "X:/AUR/25.10.2024/updated_template_excel.xlsx"
file2_path = "X:/AUR/25.10.2024/alex_report.xlsx"

def letter_to_index(letter):
    return ord(letter.upper()) - ord('A')

indexes_file1 = [letter_to_index(letter) for letter in ['A', 'B']]
indexes_file2 = [letter_to_index(letter) for letter in ['C', 'D']]

# Read the Excel files and select the specified columns by index
file1_data = pd.read_excel(file1_path, usecols=indexes_file1)
file2_data = pd.read_excel(file2_path, usecols=indexes_file2)


counter = 0
# Iterate over both DataFrames row by row
for index, (row1, row2) in enumerate(zip(file1_data.iterrows(), file2_data.iterrows())):
    line1 = row1[1]  # Get the actual data from the Series
    line2 = row2[1]

    if counter < 10:
        print(line1)
        counter += 1
    
    # Print the selected columns from both files
    print(f"Row {index + 1}:")
    for col_index1, col_index2, letter1, letter2 in zip(indexes_file1, indexes_file2, columns_file1, columns_file2):
        print(f"File 1 - Column {letter1} (Index {col_index1}): {line1[col_index1]} | "
              f"File 2 - Column {letter2} (Index {col_index2}): {line2[col_index2]}")

