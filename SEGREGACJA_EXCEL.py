#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

from finding_bom import finding_bom
from excel_check import excel
from file_segregation import file_segregation
from txt_file_creation import txt_file_creation


# pobranie linku do folderu skąd będą kopiowane wszystkie pliki:
print("\n")
print("=" * 60)
print("\n")
source = input("Podaj scieżkę do miejsca, z którego będą kopiowane pliki.\n"
               "Zostaną przeszukane wszystie foldery i podfoldery tej lokalizacji.\n ====>")
print("=" * 60)

# pobranie linku do foledru gdzie jest bom i beda kopiowane pliki:
destination = input("Podaj scieżkę do pliku BOM, tutaj zostaną skopiowane posegregowane pliki.\n"
                    " Upewnij się, że w folderze jest plik .xlsx, zawiera w nazwie'BOM'"
                    " i jest zgodny z szablonem.\n ====>")
print("=" * 60)

# nazwa pliku tekstowego z brakujacymi plikami
txt_file_name = "Lista_brakujacych_plikow.txt"

if finding_bom(destination):
    bom_path = finding_bom(destination)
    print("=" * 60)
    print(" Znaleziono plik z BOMem:\n"
          " Dokładna ścieżka do pliku z BOMem to:\n")
    print("====>   " + bom_path + "   <====")
    print("\n" + "=" * 60 + "\n")
else:
    print('bom nie został znaleziony')

#jeżeli został znaleziony plik z BOMem to otwieramy go i pobieramy dane o interesujacych nas kolumnach:
try:
    print("Pobieranie danych o wierszach i kolumnach z pliku excel.")
    part_number, tch1, tch2, tch3, rysunek, max_row = excel(bom_path)
    print("-" * 60)
except:
    print("Nie udało się pobrać danych z pliku excel, sprawdź czy zgadza się z szablonem.")

#jeżeli pobrano dane z bomu to tworzymy foldery i wrzucamy do nich pliki:
try:
    print("Segregacja plików zgodnie z obróbkami w pliku excel.")
    no_file_in_source = file_segregation(source,destination, bom_path, part_number, tch1, tch2, tch3, rysunek, max_row)
    print("-" * 60)
except:
   print("Wystapił problem w trakcie segregacji.")

#jezeli udało się posegregować pliki to wynikiem jest tabela z brakującymi elementami, poniższa funkcje wrzuca te dnae do pliku tekstowego
try:
    print("Zapis brakujących plików do pliku tekstowego: '" + txt_file_name + "'.")
    txt_file_creation(destination, no_file_in_source, txt_file_name)
    print("-" * 60)
except:
   print("Nie udało się zapisać danych o brakujących elementach do pliku.")



