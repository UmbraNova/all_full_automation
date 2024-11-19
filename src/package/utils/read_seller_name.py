import pandas as pd
import os

base_folder = "X:/21.10.2024/38-43/container 42"
file_name = "8901315_jiayou_packing list barcode.xlsx"

file_path = os.path.join(base_folder, file_name)

try:
    df = pd.read_excel(file_path, sheet_name='INVOICE', header=4)
except:
    df = pd.read_excel(file_path, sheet_name='PACKING LIST', header=4)


b8_value = df.iloc[2, 1]  # Row index 2 (third row), Column index 1 (second column)

if isinstance(b8_value, str):
    target_words = ["Address:", "Add:"]  # List of target words to search for
    words = b8_value.split()  # Split the string into words

    for target_word in target_words:
        if target_word in words:
            b8_value = ' '.join(words[:words.index(target_word)])[8:-1]
            break

print(b8_value)
