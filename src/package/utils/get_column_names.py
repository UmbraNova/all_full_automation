import pandas as pd 


excel_file_path = 'X:/AUR/23.10.2024/poze existente.xlsx'
data = pd.read_excel(excel_file_path)
    
data_top = data.head() 
    
print(data_top) 
print(list(data.columns))