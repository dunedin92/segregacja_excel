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
        rysunek = sheet.cell(row=i, column=kolumna_rysunek).value
        tch1 = sheet.cell(row=i, column=kolumna_tch1).value
        tch2 = sheet.cell(row=i, column=kolumna_tch2).value
        tch3 = sheet.cell(row=i, column=kolumna_tch3).value

        if "WYKONANY" in rysunek.upper() or ("RYSUNEK" in rysunek.upper() and "SPAWALNICZY" in rysunek.upper()):
            if tch1 != "-":
                if tch2 != "-":
                    if tch3 != "-":
                        folder_name = tch1 + "+" + tch2 + "+" + tch3
                    else:
                        folder_name = tch1 + "+" + tch2
                else:
                    folder_name = tch1

                folder_destiation = os.path.join(destination, folder_name)

                if not os.path.exists(folder_destiation):
                    os.mkdir(folder_destiation)
            else:
                no_file_in_surce.append(part_number + " - brak podanej obróbki w BOM")

            for path, dirs, files in os.walk(source):
                for format in formats:
                    part = part_number + format
                    part_destination = os.path.join(folder_destiation, part)

                    if os.path.exists(part_destination):
                        break
                    else:
                        part_source = os.path.join(path, part)
                        if os.path.exists(part_source):
                            shutil.copy(part_source, part_destination)
                            for i in no_file_in_surce:
                                if part in i:
                                    no_file_in_surce.remove(i)

                        else:
                            flag = False
                            if len(no_file_in_surce) < 1:
                                no_file_in_surce.append(part + " - " + folder_name)

                            else:
                                for i in no_file_in_surce:
                                    if part in i:
                                        flag = False
                                        break
                                    else:
                                        flag = True

                                if flag == True:
                                    no_file_in_surce.append(part + " - " + folder_name)


#

#
            #print(no_file_in_surce)
    return no_file_in_surce
