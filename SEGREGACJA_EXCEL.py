#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

from finding_bom import finding_bom
from excel_check import excel
from file_segregation import file_segregation
from txt_file_creation import txt_file_creation

# sciezka z której kopiujemy pliki:
source = 'D:\PROGRAMOWANIE\segregation_files'

# scieżka do BOM'u i miejsce gdzie będą wrzucone posegregowane pliki:
# destination = "D:\PROGRAMOWANIE\pliki"

# pobranie linku do foledru gdzie jest bom i beda kopiowane pliki:
destination = input("Podaj scieżkę do pliku BOM, tutaj zostaną skopiowane posegregowane pliki.\n"
                    " Upewnij się, że w folderze jest plik .xlsx i jest to BOM.\n ==>")
txt_file_name = "Lista_brakujących_plików.txt"

if finding_bom(destination):
    bom_path = finding_bom(destination)
    print("=" * 60)
    print(" Znaleziono plik z BOMem:\n"
          " Dokładna ścieżka do pliku z BOMem to:\n")
    print("====>   " + bom_path + "   <====")
    print("\n" + "=" * 60 + "\n")

#jeżeli został znaleziony plik z BOMem to otwieramy go i pobieramy dane o interesujacych nas kolumnach:
#    try:
part_number, tch1, tch2, tch3, rysunek, max_row = excel(bom_path)
#    except:
#        print("Nie udało się pobrać danych z pliku excel, sprawdź czy zgadza się z szablonem")

#jeżeli pobrano dane z bomu to tworzymy foldery i wrzucamy do nich pliki:
#    try:
no_file_in_source = file_segregation(source,destination, bom_path, part_number, tch1, tch2, tch3, rysunek, max_row)
print(no_file_in_source)
#    except:
 #       print("wystapił problem w trakcie segregacji")

#jezeli udało się posegregować pliki to wynikiem jest tabela z brakującymi elementami, poniższa funkcje wrzuca te dnae do pliku tekstowego
#    try:
txt_file_creation(destination, no_file_in_source, txt_file_name)
 #   except:
 #       print("Nie udało się zapisać danych o brakujących elementach do pliku")

#else:
#    print("nie znaleziono pliku excel z BOMem")
#    exit()

