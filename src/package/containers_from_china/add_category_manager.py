# def main(template_file_path="X:/AUR/28.10.2024/CONTAINER 44 DE REFACUT.xlsx", categories_with_cm_path="X:/AUR/28.10.2024/CATEGORIES WITH CM.xlsx"):
def main(template_file_path="X:/AUR/11.2024/08.11.2024/noul template.xlsx", categories_with_cm_path="X:/AUR/11.2024/08.11.2024/CATEGORIES WITH CM.xlsx"):
    import pandas as pd


    file1_path = template_file_path
    file2_path = categories_with_cm_path

    def letter_to_index(letter):
        return ord(letter.upper()) - ord('A')

    indexes_file1 = [letter_to_index(letter) for letter in ['N', 'P']]
    indexes_file2 = [letter_to_index(letter) for letter in ['G', 'H']]

    # Read the Excel files and select the specified columns by index
    file1_data = pd.read_excel(file1_path, usecols=indexes_file1)
    file2_data = pd.read_excel(file2_path, usecols=indexes_file2)

    cm_codes_dict = {}

    # Iterate over DataFrame row by row
    def dataframe_create_dict(file_data):
        for i, row in file_data.iterrows():
            retail_code = str(row.iloc[0])
            cm_name = str(row.iloc[1])

            if retail_code in cm_codes_dict:
                cm_codes_dict[retail_code].append(cm_name)
            else:
                cm_codes_dict[retail_code] = [cm_name]


    def compare_dataframe_with_dict(file_data):
        # Get the position of the column corresponding to "P" within file_data
        target_column_name = file_data.columns[1]
        target_column_position = file_data.columns.get_loc(target_column_name)

        for i, row in file_data.iterrows():
            retail_code = str(row.iloc[0])
            if retail_code in cm_codes_dict:
                cm_name = cm_codes_dict[retail_code][0]
                print(f"Updating row {i}, column '{target_column_name}': {retail_code} -> {cm_name}")
                file_data.iat[i, target_column_position] = cm_name  # Update the specific cell by column position

        print("Completed!")


    dataframe_create_dict(file2_data)
    compare_dataframe_with_dict(file1_data)


if __name__ == "__main__":
    main()