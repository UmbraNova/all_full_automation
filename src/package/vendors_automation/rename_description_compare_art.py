def main(template_file_path="X:/AUR/11.2024/01.11.2024/noul template - m.xlsx"):
# def main(template_file_path="X:/AUR/21.10.2024 - Copy/noul template.xlsx"):
    import openpyxl
    import re


    # redenumeste toate produsele, le curata de unele elemente si innlocuieste cu altele
    # verifica daca codul de articol coincide cu cel din coloana cu coduri de articole(acelasi file)

    def change_value(value):
        # replacing_el = ["#", ".", "Ă", "Â", "Î", "Ș", "Ț", "PR.", "*", "APT.", " COD ", "COL.COD", "ARTICOLUL", "COD PR.", ".ART.ART.", "ART.ART.", "ART ART.", "ART ART", " ART ART ", "IN STOC", "ART.PR", "ART PR", "UN SET", "IN SORTIMENT", " ,", "#", '"', "NR ", "ART ", "SETUL", " PR ", " SORT ", "FUNCTIONEAZA", "PENTRU", "PT."]
        replacing_el = ["-", '"', "  ", "PENTRU", "PT.", "DE"]
        replacing_with = [" ", "", " ", "PT", "PT", ""]

        for i in range(len(replacing_el)):
            value = str(value).upper()
            value = re.sub(r'\s+', " ", value)
            value = value.replace(replacing_el[i], replacing_with[i])
        return value

    # //////////////////////////////////////////////////////////////////// check&change pentru numar aticol
    def check_number_at_end(value, number):
        number_str = str(number)
        
        if value.endswith(number_str):
            return value
        else:
            return f"{value}"


    def get_column_value(sheet, row):
        column = "O"  # column to check info from
        cell_value = sheet[f"{column}{row}"].value
        return cell_value
    # //////////////////////////////////////////////////////////////////// check&change pentru numar aticol



    def process_excel(file_path, columns):
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active  # Select the active sheet

        for col in columns:
            for row in range(1, sheet.max_row + 1):  # Starting from row 1
                cell = sheet[f"{col}{row}"]  # Access the cell
                if cell.value is not None:  # Check if the cell is not empty
    # //////////////////////////////////////////////////////////////////// check&change pentru numar aticol
                    item_code = get_column_value(sheet, row)
                    cell.value = check_number_at_end(cell.value, item_code)
    # //////////////////////////////////////////////////////////////////// check&change pentru numar aticol
                    cell.value = change_value(cell.value)  # Change the cell value


        print("Values changed")
        workbook.save(file_path)

    file_path = template_file_path
    process_excel(file_path, ["C"])


if __name__ == "__main__":
    main()
    