import os
import shutil

def move_files(destination, no_file_in_surce):

    file_formats = [".step", ".stl", ".pdf", ".dxf"]

    plik = open("temp_file_txt.txt", "r", encoding='utf8')

    list = plik.readlines()

    for line in list:
        if line.upper() != "OK":
            path, folder = line.split(" => ")
            path = path.rstrip()
            folder = folder.rstrip()
            print(path)
            print(folder)

            for format in file_formats:
                path1 = path + format
                name = os.path.basename(path1)
                print("pobrana nazwa pliku to:  " + name)

                if os.path.exists(path1):
                    destination_path = os.path.join(destination, folder)
                    destination_path = os.path.join(destination_path, name)

                    print("kopiujemy sciezke: " + path1)
                    print("kopiujemy siezke do: " + destination_path)
                    shutil.move(path1, destination_path)
                else:
                    name_to_write = name + "  " + folder

                    if name_to_write in no_file_in_surce:
                        print("brak pliku został juz odnowotowany")
                    else:
                        no_file_in_surce.append(name + "  " + folder)

    return no_file_in_surce


