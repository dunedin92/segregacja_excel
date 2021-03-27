# Funkcja przeszukująca excel i tworząca fodlery zgodnie z kolumnami tch1 tch2 tch3 rysunek
# Przegladajaca excel i przenoszaca pliki do odpowiednich folderów
# W zależności od folderów kopiuje odpowiednie pliki (do C pdf i dxf, do F P pdf i step, do DRUK 3D stl, itp)
# Tworzy plik tekstowy z brakującymi pliki i dopisuje do jakiego folderu powinien trafic brakujacy plik
import os
import shutil
import openpyxl


def file_segregation(source, destination, bom_path,  kolumna_part_number, kolumna_tch1, kolumna_tch2, kolumna_tch3,
                     kolumna_rysunek, max_row):
    wb = openpyxl.load_workbook(bom_path)
    type(wb)
    arkusze = wb.sheetnames
    sheet = wb[arkusze[0]]

    formats = [".pdf", ".dxf", ".step", ".stl"]
    no_file_in_surce = []

    for i in range(2, max_row + 1):

        part_number = sheet.cell(row=i, column=kolumna_part_number).value
        part_number = part_number.lstrip()
        print(part_number)
        rysunek = sheet.cell(row=i, column=kolumna_rysunek).value
        tch1 = sheet.cell(row=i, column=kolumna_tch1).value
        tch2 = sheet.cell(row=i, column=kolumna_tch2).value
        tch3 = sheet.cell(row=i, column=kolumna_tch3).value

        if "WYKONANY" in rysunek.upper() or ("RYSUNEK" in rysunek.upper() and "SPAWALNICZY" in rysunek.upper()):
            print('Istnieje rysunek do tej części')
            if tch1 != "-":
                if tch2 != "-":
                    if tch3 != "-":
                        folder_name = tch1 + "+" + tch2 + "+" + tch3
                    else:
                        folder_name = tch1 + "+" + tch2
                else:
                    folder_name = tch1

                folder_destiation = os.path.join(destination, folder_name)
                print(folder_destiation)

                if os.path.exists(folder_destiation):
                    print('folder został juz utworzony')
                else:
                    os.mkdir(folder_destiation)
            else:
                print(part_number)
                print("dla tego pliku nie ma przypisanej obróbki")
                no_file_in_surce.append(part_number + "brak podanej obróbki w BOM")

            for format in formats:
                part = part_number + format
                print("-" * 60)
                print(part)
                part_destination = os.path.join(folder_destiation, part)

                for path, dirs, files in os.walk(source):
                    print(path)
                    part_source = os.path.join(path, part)
                    if os.path.exists(part_source):
                        if os.path.exists(part_destination):
                            print("taki plik został już wczesniej przeniesiony")
                            break
                        else:
                            shutil.copy(part_source, part_destination)
                            break
                    else:
                        if part in no_file_in_surce:
                            print("brak tego pliku został już odnotowany")
                        else:
                            print('doddajemy rekord do tablicy')
                            no_file_in_surce.append(part + " - " + folder_name)

                    print(no_file_in_surce)
    return no_file_in_surce
