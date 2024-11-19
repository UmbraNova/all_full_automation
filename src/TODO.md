the trecut prin algoritmul de schimbare a numelui fiecare .zip (dezarhivat mai intai) cu ITM293932

apoi de arhivat :T_T si de trecut prin algoritmul de stergere
-------------------------------------------------------------------

- Folderele sunt arhivate, au poze denumite conform cu 041024.xlsx.
- De redenumit cu numarul ITM000941 si sterse duplicates if exists.
- Apoi vom avea un folder dezarhivat cu poze denumite 70446.jpg in loc de 70446-2-3.jpg

apoi el a extras din system lista cu produse care deja exista (denumirile ITM000412)
le pune in acel excel din arhiva 041024.xlsx

apoi eu le redenumesc din 70446.jpg in ITM000412.jpg conform tabelului 041024.xlsx actualizat

apoi conform excelului poze existente.xlsx compar pozele redenumite cu cele din coloanal A si daca in coloanal B scrie No sau in coloanal C e empty value, le sterg

YES/NO = fara poza / cu poza

ITM YES 00 fara poza
ITM YES __ fara poza
ITM YES 00 fara poza
ITM NO 00 fara poza

ITM YES 00 cu poza
ITM YES 00 cu poza
ITM YES 00 cu poza
ITM NO 00 cu poza
ITM NO __ cu poza
ITM NO 00 cu poza
ITM NO __ cu poza
ITM YES __ cu poza


-------------------------------------------------------------------- varianta actuala:
- dezarhivat foldere
- combinat intr-un folder mare (combinat excelurile aferente intr-un excel mare)?
- redenumit din 70446-2-3.jpg in 70446.jpg, cu renamte_to_numbers.py conform cu excelul din fiecare
- de extras din sistem lista cu produse care deja exista (denumirile ITM000412)
- le pune cu VLOOKUP in lista articol_conform_ITM.xlsx
- redenumesc din 70446.jpg in ITM000412 cu rename_to_itm000.py conform articol_conform_ITM.xlsx
- 


-------------------------------------------------------------------- inregistrare teams rezumat:

in Business Central > search > Retail image list > (avem nume de poze aleatorii)
(ar fi cel mai bine sa primim numele produsului in denumirea imaginii)
Razvan - "O sa vrem sa facem un job la un moment dat, sa curete pozele fara legatura cu un produs"
(Configuration Packages - afecteaza performanta sitemului?)
Poti sa importi in sistem maxim 28 MB
Retail image list > import from Zip
(sa nu aiba nici o poza mai multa de 500 KB)
(de verificat daca sunt in sistem)
(daca ii dai export de sus iti scoate toate tabelurile, daca dai export de jos iti scoate doar unul)
in Configuration Packages > ITEM PICTURES > Excel > Export to Excel > (Image Id e nume poza/nume produs)

si iti scoate un Default24_10_2024_12_54_19.xlsx, din acest default comparam coloana A si Z
adica No.: ITM000001 si Attrib 1 Code: 76436  (le schimb cu locul)
in excelul din folderul arhivat cu poze, scriu =VLOOKUP(Код товару, si cele 2 coloane din Default24...)
pentru a vedea ce Attrib 1 Code are deja nume produs ITM000001
dupa ce am codul de ITM (daca il am), CTRL + C / CTRL + SHIFT + V

-------------------------------------------------------------------- din main.py:



'''for containers from china'''

# containers_excel_automation.py
# vendors_name_to_code.py
# art_compare_barcode.py
# rename_description_compare_art.py
# compare_columns_diff_files.py

# corectat description col C
# adaugat Green Tax True or False col S / si amount col T
# adaugat numar bucati col U
# automatizat nume brand sa fie pus primul

# description_text1_text2.py


# attrib 4 code, category manager de pus (VLOOKUP deocamdata)

'''for image processing'''

# the trecut prin algoritmul de schimbare a numelui fiecare .zip (dezarhivat mai intai) cu ITM293932

# apoi de arhivat :T_T si de trecut prin algoritmul de stergere
# -------------------------------------------------------------------

# - Folderele sunt arhivate, au poze denumite conform cu 041024.xlsx.
# - De redenumit cu numarul ITM000941 si sterse duplicates if exists.
# - Apoi vom avea un folder dezarhivat cu poze denumite 70446.jpg in loc de 70446-2-3.jpg

# apoi el a extras din system lista cu produse care deja exista (denumirile ITM000412)
# le pune in acel excel din arhiva 041024.xlsx

# apoi eu le redenumesc din 70446.jpg in ITM000412.jpg conform tabelului 041024.xlsx actualizat

# apoi conform excelului poze existente.xlsx compar pozele redenumite cu cele din coloanal A si daca in coloanal B scrie No sau in coloanal C e empty value, le sterg

# YES/NO = fara poza / cu poza

# ITM YES 00 fara poza
# ITM YES __ fara poza
# ITM YES 00 fara poza
# ITM NO 00 fara poza

# ITM YES 00 cu poza
# ITM YES 00 cu poza
# ITM YES 00 cu poza
# ITM NO 00 cu poza
# ITM NO __ cu poza
# ITM NO 00 cu poza
# ITM NO __ cu poza
# ITM YES __ cu poza


# -------------------------------------------------------------------- varianta actuala:
# - dezarhivat foldere
# - combinat intr-un folder mare (combinat excelurile aferente intr-un excel mare)?
# - redenumit din 70446-2-3.jpg in 70446.jpg, cu renamte_to_numbers.py conform cu excelul din fiecare
# - de extras din sistem lista cu produse care deja exista (denumirile ITM000412)
# - le pune cu VLOOKUP in lista articol_conform_ITM.xlsx
# - redenumesc din 70446.jpg in ITM000412 cu rename_to_itm000.py conform articol_conform_ITM.xlsx
# - 


# -------------------------------------------------------------------- inregistrare teams rezumat:

# in Business Central > search > Retail image list > (avem nume de poze aleatorii)
# (ar fi cel mai bine sa primim numele produsului in denumirea imaginii)
# Razvan - "O sa vrem sa facem un job la un moment dat, sa curete pozele fara legatura cu un produs"
# (Configuration Packages - afecteaza performanta sitemului?)
# Poti sa importi in sistem maxim 28 MB
# Retail image list > import from Zip
# (sa nu aiba nici o poza mai multa de 500 KB)
# (de verificat daca sunt in sistem)
# (daca ii dai export de sus iti scoate toate tabelurile, daca dai export de jos iti scoate doar unul)
# in Configuration Packages > ITEM PICTURES > Excel > Export to Excel > (Image Id e nume poza/nume produs)

# si iti scoate un Default24_10_2024_12_54_19.xlsx, din acest default comparam coloana A si Z
# adica No.: ITM000001 si Attrib 1 Code: 76436  (le schimb cu locul)
# in excelul din folderul arhivat cu poze, scriu =VLOOKUP(Код товару, si cele 2 coloane din Default24...)
# pentru a vedea ce Attrib 1 Code are deja nume produs ITM000001
# dupa ce am codul de ITM (daca il am), CTRL + C / CTRL + SHIFT + V


# deci, trebuie de facut:

# upload la imagini, dar doar la cele care nu sunt in sistem.

# de asta le transcriu din random number in exact numbers si dupa in ITM.
# dupa le compar denumirea de ITM cu ITM existente in alex_raport.xlsx






































