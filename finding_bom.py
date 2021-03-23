# Przeszukiwanie folderu w celu znalezienia BOM'u.
# Plik musi mieć rozszerzenie .xlsx, i zawierać w nazwie człon "BOM" (wielkość znaków nieistotna)
import os


def finding_bom(destination):
    pliki_excela = []
    for path, dirs, files in os.walk(destination):
        for name in files[0:]:
            if name[-4:] == "xlsx":
                pliki_excela.append(name)

    if len(pliki_excela) >= 1:
        for nazwa in pliki_excela:
            if "BOM" in nazwa.upper():
                bom = nazwa
                bom_path = os.path.join(destination, bom)
                return bom_path
            else:
                return False
    else:
        return False
