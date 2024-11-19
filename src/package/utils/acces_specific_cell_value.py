import pandas as pd

df = pd.read_excel('X:\\21.10.2024\\38-43\\CONTAINER  41\\8901296_changyou_invoice barcode.xlsx', 
                   sheet_name='INVOICE', header=4)

# Access the specific cell value directly
desired_value = df.iloc[2, 1]  # Row index 2 (third row), Column index 1 (second column)

print(desired_value)
