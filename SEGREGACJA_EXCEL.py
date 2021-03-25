#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

from finding_bom import finding_bom
from excel_check import excel
from file_segregation import file_segregation

# sciezka z której kopiujemy pliki:
source = 'D:\PROGRAMOWANIE\segregation_files'

# scieżka do BOM'u i miejsce gdzie będą wrzucone posegregowane pliki:
# destination = "D:\PROGRAMOWANIE\pliki"

# pobranie linku do foledru gdzie jest bom i beda kopiowane pliki:
destination = input("Podaj scieżkę do pliku BOM, tutaj zostaną skopiowane posegregowane pliki.\n"
                    " Upewnij się, że w folderze jest plik .xlsx i jest to BOM.\n ==>")

if finding_bom(destination):
    bom_path = finding_bom(destination)
    print("=" * 60)
    print(" Znaleziono plik z BOMem:\n"
          " Dokładna ścieżka do pliku z BOMem to:\n")
    print("====>   " + bom_path + "   <====")
    print("\n" + "=" * 60 + "\n")

#jeżeli został znaleziony plik z BOMem to otwieramy go i pobieramy dane o interesujacych nas kolumnach:
    part_number, tch1, tch2, tch3, rysunek, max_row = excel(bom_path)



else:
    print("nie znaleziono pliku excel z BOMem")
    exit()

