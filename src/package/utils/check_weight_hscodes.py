import pandas as pd


file_path = "X:/AUR/11.2024/08.11.2024/total_import_noul template.xlsx"
df = pd.read_excel(file_path, header=None)
df[3] = pd.to_numeric(df[3], errors='coerce')
df[4] = pd.to_numeric(df[4], errors='coerce')

def check_weight(row):
    # if row.iloc[2] > row.iloc[1]:
    if row[4] > row[3]:
        return "Weight issues"
    return ""

def check_codes(row):
    if len(str(row[5])) > 8 or len(str(row[5])) < 8 or check_string(str(row[5])):
        return "HS code issues"
    return ""

def check_string(data):
    # print(data)
    has_numbers = any(char.isdigit() for char in data)
    has_letters = any(char.isalpha() for char in data)
    has_symbols = any(not char.isalnum() for char in data)

    # print("Contains numbers:", has_numbers)
    # print("Contains letters:", has_letters)
    # print("Contains symbols:", has_symbols)

    if has_numbers:
        if  has_letters or has_symbols:
            # print(has_letters, has_symbols, "<-=====")
            return True
        return False

df[6] = df.apply(check_weight, axis=1)
df[7] = df.apply(check_codes, axis=1)
df.to_excel(file_path, header=False, index=False)
print("Cleaning complete. The cleaned file is saved as:", file_path)
