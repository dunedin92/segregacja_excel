# Otworzenie EXCELa, otworzenie pierwszego arkusza;
# - znalezienie max ilości wierszy
# - znalezienie kolumn: part number, rysunek, i obróbek
# - zwrócenie wszystkich tych wartości
import openpyxl


def excel(path):
    global kolumna_tch1, kolumna_part_number, kolumna_tch2, kolumna_tch3, kolumna_rysunek
    wb = openpyxl.load_workbook(path)
    type(wb)
    arkusze = wb.sheetnames
    sheet = wb[arkusze[0]]

    max_column = sheet.max_column
    max_row = sheet.max_row

    for i in range(1, max_column + 1):

        value = sheet.cell(row=1, column=i).value
        if "TCH" in value.upper() and "1:" in value.upper():
            kolumna_tch1 = i

        if "TCH" in value.upper() and "2:" in value.upper():
            kolumna_tch2 = i

        if "TCH" in value.upper() and "3:" in value.upper():
            kolumna_tch3 = i

        if "RYSUNEK" in value.upper():
            kolumna_rysunek = i

        if "PART" in value.upper() and "NUMBER" in value.upper():
            kolumna_part_number = i

    #  print(kolumna_part_number)
    #  print(kolumna_tch1)
    #  print(kolumna_tch2)
    #  print(kolumna_tch3)
    #  print(kolumna_rysunek)
    return kolumna_part_number, kolumna_tch1, kolumna_tch2, kolumna_tch3, kolumna_rysunek, max_row
