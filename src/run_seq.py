import package.containers_from_china.containers_excel_automation as containers_excel_automation
import package.containers_from_china.add_category_manager as add_category_manager
import package.containers_from_china.art_compare_barcode as art_compare_barcode
import package.containers_from_china.compare_columns_diff_files as compare_columns_diff_files
import package.containers_from_china.description_text1_text2 as description_text1_text2
import package.containers_from_china.rename_description_compare_art as rename_description_compare_art
import package.containers_from_china.vendors_name_to_code as vendors_name_to_code

template_file_path = "X:/AUR/containers_auto/noul template.xlsx"
containers_folder_path = "X:/AUR/containers_auto/containers_folder"
categories_with_cm_path = "X:/AUR/containers_auto/categories_with_cm/CATEGORIES WITH CM.xlsx"
items_in_system_path = "X:/AUR/containers_auto/items_in_system/Default19_11_2024_16_16_41.xlsx"  # TODO: change it
vendors_file_path = "X:/AUR/containers_auto/vendors_file/Vendors.xlsx"


def run_seq_1():
    containers_excel_automation.main(template_file_path, containers_folder_path)
    vendors_name_to_code.main(template_file_path, vendors_file_path)
    art_compare_barcode.main(template_file_path)
    rename_description_compare_art.main(template_file_path)
    add_category_manager.main(template_file_path, categories_with_cm_path)
    compare_columns_diff_files.main(template_file_path, items_in_system_path)  # compara cu items din sistem | Configuration Packages > ITEM > Export to Excel


def run_seq_2():
    description_text1_text2.main(template_file_path)


def process_files():
    run_seq_1()
    # de verificat description, corectat formularea unde e cazul col C
    # adaugat Green Tax True or False col S / si amount col T
    # adaugat numar bucati col U
    # TODO: automatizat nume brand sa fie pus primul
    run_seq_2()

process_files()

