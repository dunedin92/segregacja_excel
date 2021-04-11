#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QMainWindow, QTextEdit,
                             QAction, QFileDialog, QApplication)
from PyQt5.QtGui import QIcon
import sys
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QLineEdit
import sys

class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(200, 200, 640, 480)
        self.setWindowTitle('Okienko')
        self.initUI()

    def initUI(self):
        self.text1 = QtWidgets.QLabel(self)
        self.text1.setText('Wybierz folder, z którego będą kopiowane pliki:')
        self.text1.move(10, 10)
        self.text1.adjustSize()

        self.link1 = QLineEdit(self)
        self.link1.resize(390, 24)
        self.link1.setText('Wpisz link do folderu lub kliknij: Przeglądaj')
        self.link1.move(10, 30)

        self.source_button = QtWidgets.QPushButton(self)
        self.source_button.setText('Przeglądaj')
        self.source_button.adjustSize()
        self.source_button.clicked.connect(self.clicked1)
        self.source_button.move(400, 30)

        self.text2 = QtWidgets.QLabel(self)
        self.text2.setText('Wybierz folder, do którego będą kopiowane posegregowane pliki:')
        self.text2.move(10, 80)
        self.text2.adjustSize()

        self.link2 = QLineEdit(self)
        self.link2.resize(390, 24)
        self.link2.setText('Wpisz link do folderu lub kliknij: Przeglądaj')
        self.link2.move(10, 110)

        self.destination_button = QtWidgets.QPushButton(self)
        self.destination_button.setText('Przeglądaj')
        self.destination_button.adjustSize()
        self.destination_button.clicked.connect(self.clicked2)
        self.destination_button.move(400, 110)

        self.text3 = QtWidgets.QLabel(self)
        self.text3.setText('Wybierz plik programu EXCEL zawierający BOM:')
        self.text3.move(10, 160)
        self.text3.adjustSize()



        self.bom_button = QtWidgets.QPushButton(self)
        self.bom_button.setText('Przeglądaj')
        self.bom_button.adjustSize()
        self.bom_button.clicked.connect(self.clicked3)
        self.bom_button.move(50, 180)



    def clicked1(self):
        self.source = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.link1.setText(self.source)

    def clicked2(self):
        self.destination = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.link2.setText(self.destination)

    def clicked3(self):
        self.bom_path = QFileDialog.ExistingFile(self, "Select File Name:", "\D",  " Excel files (*.xlsx *.xls)")
        self.text3.setText(self.bom_path)



def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())

window()





# # pobranie linku do folderu skąd będą kopiowane wszystkie pliki:
# print("\n")
# print("=" * 60)
# print("\n")
# source = input("Podaj scieżkę do miejsca, z którego będą kopiowane pliki.\n"
#                "Zostaną przeszukane wszystie foldery i podfoldery tej lokalizacji.\n ====>")
# print("=" * 60)
#
# # pobranie linku do foledru gdzie jest bom i beda kopiowane pliki:
# destination = input("Podaj scieżkę do pliku BOM, tutaj zostaną skopiowane posegregowane pliki.\n"
#                     " Upewnij się, że w folderze jest plik .xlsx, zawiera w nazwie'BOM'"
#                     " i jest zgodny z szablonem.\n ====>")
# print("=" * 60)
#
# # nazwa pliku tekstowego z brakujacymi plikami
# txt_file_name = "Lista_brakujacych_plikow.txt"
#
# if finding_bom(destination):
#     bom_path = finding_bom(destination)
#     print("=" * 60)
#     print(" Znaleziono plik z BOMem:\n"
#           " Dokładna ścieżka do pliku z BOMem to:\n")
#     print("====>   " + bom_path + "   <====")
#     print("\n" + "=" * 60 + "\n")
# else:
#     print('bom nie został znaleziony')
#
# #jeżeli został znaleziony plik z BOMem to otwieramy go i pobieramy dane o interesujacych nas kolumnach:
# try:
#     print("Pobieranie danych o wierszach i kolumnach z pliku excel.")
#     part_number, tch1, tch2, tch3, rysunek, max_row = excel(bom_path)
#     print("-" * 60)
# except:
#     print("Nie udało się pobrać danych z pliku excel, sprawdź czy zgadza się z szablonem.")
#
# #jeżeli pobrano dane z bomu to tworzymy foldery i wrzucamy do nich pliki:
# try:
#     print("Segregacja plików zgodnie z obróbkami w pliku excel.")
#     no_file_in_source = file_segregation(source,destination, bom_path, part_number, tch1, tch2, tch3, rysunek, max_row)
#     print("-" * 60)
# except:
#    print("Wystapił problem w trakcie segregacji.")
#
# #jezeli udało się posegregować pliki to wynikiem jest tabela z brakującymi elementami, poniższa funkcje wrzuca te dnae do pliku tekstowego
# try:
#     print("Zapis brakujących plików do pliku tekstowego: '" + txt_file_name + "'.")
#     txt_file_creation(destination, no_file_in_source, txt_file_name)
#     print("-" * 60)
# except:
#    print("Nie udało się zapisać danych o brakujących elementach do pliku.")