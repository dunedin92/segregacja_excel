#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from excel_check import excel_check
from qty_total_calculation import qty_total_calculation
from txt_file_creation import txt_file_creation
from temp_file_list import temp_file_list
from move_files import move_files
import os
import sys
import webbrowser
import time
import subprocess
import shutil


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setObjectName("MainWindow")
        self.resize(640, 720)
        self.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.initUI()

    def initUI(self):
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.block_1 = QtWidgets.QFrame(self.centralwidget)
        self.block_1.setGeometry(QtCore.QRect(10, 10, 620, 155))
        self.block_1.setAccessibleName("")
        self.block_1.setAccessibleDescription("")
        self.block_1.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.block_1.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_1.setFrameShadow(QtWidgets.QFrame.Plain)
        self.block_1.setLineWidth(2)
        self.block_1.setObjectName("block_1")
        self.block_1_title = QtWidgets.QLabel(self.block_1)
        self.block_1_title.setGeometry(QtCore.QRect(10, 5, 200, 20))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(18)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.block_1_title.setFont(font)
        self.block_1_title.setObjectName("block_1_title")
        self.block_1_description = QtWidgets.QLabel(self.block_1)
        self.block_1_description.setGeometry(QtCore.QRect(10, 35, 600, 30))
        self.block_1_description.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.block_1_description.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_1_description.setFrameShadow(QtWidgets.QFrame.Raised)
        self.block_1_description.setObjectName("block_1_description")
        self.bom_button = QtWidgets.QPushButton(self.block_1)
        self.bom_button.setGeometry(QtCore.QRect(10, 80, 80, 25))
        self.bom_button.setObjectName("bom_button")
        self.bom_verification_button = QtWidgets.QPushButton(self.block_1)
        self.bom_verification_button.setGeometry(QtCore.QRect(10, 120, 200, 25))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.bom_verification_button.setFont(font)
        self.bom_verification_button.setDefault(True)
        self.bom_verification_button.setObjectName("bom_verification_button")
        self.line_bom_path = QtWidgets.QLineEdit(self.block_1)
        self.line_bom_path.setGeometry(QtCore.QRect(90, 80, 520, 25))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setItalic(True)
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.line_bom_path.setFont(font)
        self.line_bom_path.setInputMask("")
        self.line_bom_path.setText("Nie wybrano pliku.")
        self.line_bom_path.setMaxLength(32773)
        self.line_bom_path.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.line_bom_path.setReadOnly(True)
        self.line_bom_path.setObjectName("line_bom_path")
        self.bom_verification_info_text = QtWidgets.QLabel(self.block_1)
        self.bom_verification_info_text.setGeometry(QtCore.QRect(220, 120, 390, 25))
        self.bom_verification_info_text.setObjectName("bom_verification_info_text")
        self.block_2 = QtWidgets.QFrame(self.centralwidget)
        self.block_2.setGeometry(QtCore.QRect(10, 170, 620, 120))
        self.block_2.setAccessibleName("")
        self.block_2.setAccessibleDescription("")
        self.block_2.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.block_2.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_2.setFrameShadow(QtWidgets.QFrame.Plain)
        self.block_2.setLineWidth(2)
        self.block_2.setObjectName("block_2")
        self.block_2_title = QtWidgets.QLabel(self.block_2)
        self.block_2_title.setGeometry(QtCore.QRect(10, 5, 200, 20))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(18)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.block_2_title.setFont(font)
        self.block_2_title.setObjectName("block_2_title")
        self.block_2_description = QtWidgets.QLabel(self.block_2)
        self.block_2_description.setGeometry(QtCore.QRect(10, 35, 600, 40))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(2)
        sizePolicy.setHeightForWidth(self.block_2_description.sizePolicy().hasHeightForWidth())
        self.block_2_description.setSizePolicy(sizePolicy)
        self.block_2_description.setMinimumSize(QtCore.QSize(0, 30))
        self.block_2_description.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.block_2_description.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_2_description.setFrameShadow(QtWidgets.QFrame.Raised)
        self.block_2_description.setTextFormat(QtCore.Qt.AutoText)
        self.block_2_description.setScaledContents(False)
        self.block_2_description.setWordWrap(True)
        self.block_2_description.setObjectName("block_2_description")
        self.qty_calculation_button = QtWidgets.QPushButton(self.block_2)
        self.qty_calculation_button.setGeometry(QtCore.QRect(10, 85, 200, 23))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.qty_calculation_button.setFont(font)
        self.qty_calculation_button.setDefault(True)
        self.qty_calculation_button.setObjectName("qty_calculation_button")
        self.qty_calculation_status = QtWidgets.QLabel(self.block_2)
        self.qty_calculation_status.setGeometry(QtCore.QRect(220, 85, 390, 25))
        self.qty_calculation_status.setObjectName("qty_calculation_status")
        self.block_3 = QtWidgets.QFrame(self.centralwidget)
        self.block_3.setGeometry(QtCore.QRect(10, 295, 620, 120))
        self.block_3.setAccessibleName("")
        self.block_3.setAccessibleDescription("")
        self.block_3.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.block_3.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_3.setFrameShadow(QtWidgets.QFrame.Plain)
        self.block_3.setLineWidth(2)
        self.block_3.setObjectName("block_3")
        self.block_3_title = QtWidgets.QLabel(self.block_3)
        self.block_3_title.setGeometry(QtCore.QRect(10, 5, 200, 20))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(18)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.block_3_title.setFont(font)
        self.block_3_title.setObjectName("block_3_title")
        self.block_3_description = QtWidgets.QLabel(self.block_3)
        self.block_3_description.setGeometry(QtCore.QRect(10, 35, 600, 30))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.block_3_description.sizePolicy().hasHeightForWidth())
        self.block_3_description.setSizePolicy(sizePolicy)
        self.block_3_description.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.block_3_description.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_3_description.setFrameShadow(QtWidgets.QFrame.Raised)
        self.block_3_description.setObjectName("block_3_description")
        self.button_destination_path = QtWidgets.QPushButton(self.block_3)
        self.button_destination_path.setGeometry(QtCore.QRect(530, 80, 80, 25))
        self.button_destination_path.setDefault(True)
        self.button_destination_path.setObjectName("button_destination_path")
        self.line_destination_path = QtWidgets.QLineEdit(self.block_3)
        self.line_destination_path.setGeometry(QtCore.QRect(10, 80, 520, 25))
        self.line_destination_path.setInputMask("")
        self.line_destination_path.setReadOnly(False)
        self.line_destination_path.setObjectName("line_destination_path")
        self.block_4 = QtWidgets.QFrame(self.centralwidget)
        self.block_4.setGeometry(QtCore.QRect(10, 420, 620, 120))
        self.block_4.setAccessibleName("")
        self.block_4.setAccessibleDescription("")
        self.block_4.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.block_4.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_4.setFrameShadow(QtWidgets.QFrame.Plain)
        self.block_4.setLineWidth(2)
        self.block_4.setObjectName("block_4")
        self.block_4_title = QtWidgets.QLabel(self.block_4)
        self.block_4_title.setGeometry(QtCore.QRect(10, 5, 200, 20))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(18)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.block_4_title.setFont(font)
        self.block_4_title.setObjectName("block_4_title")
        self.block_4_description = QtWidgets.QLabel(self.block_4)
        self.block_4_description.setGeometry(QtCore.QRect(10, 35, 600, 30))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.block_4_description.sizePolicy().hasHeightForWidth())
        self.block_4_description.setSizePolicy(sizePolicy)
        self.block_4_description.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.block_4_description.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_4_description.setFrameShadow(QtWidgets.QFrame.Raised)
        self.block_4_description.setObjectName("block_4_description")
        self.button_source_path = QtWidgets.QPushButton(self.block_4)
        self.button_source_path.setGeometry(QtCore.QRect(530, 80, 80, 25))
        self.button_source_path.setDefault(True)
        self.button_source_path.setObjectName("button_source_path")
        self.line_source_path = QtWidgets.QLineEdit(self.block_4)
        self.line_source_path.setGeometry(QtCore.QRect(10, 80, 520, 25))
        self.line_source_path.setInputMask("")
        self.line_source_path.setReadOnly(False)
        self.line_source_path.setObjectName("line_source_path")
        self.block_5 = QtWidgets.QFrame(self.centralwidget)
        self.block_5.setGeometry(QtCore.QRect(10, 545, 620, 120))
        self.block_5.setAccessibleName("")
        self.block_5.setAccessibleDescription("")
        self.block_5.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.block_5.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_5.setFrameShadow(QtWidgets.QFrame.Plain)
        self.block_5.setLineWidth(2)
        self.block_5.setObjectName("block_5")
        self.block_5_title = QtWidgets.QLabel(self.block_5)
        self.block_5_title.setGeometry(QtCore.QRect(10, 5, 200, 20))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(18)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.block_5_title.setFont(font)
        self.block_5_title.setObjectName("block_5_title")
        self.block_5_description = QtWidgets.QLabel(self.block_5)
        self.block_5_description.setGeometry(QtCore.QRect(10, 35, 600, 30))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.block_5_description.sizePolicy().hasHeightForWidth())
        self.block_5_description.setSizePolicy(sizePolicy)
        self.block_5_description.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.block_5_description.setFrameShape(QtWidgets.QFrame.Panel)
        self.block_5_description.setFrameShadow(QtWidgets.QFrame.Raised)
        self.block_5_description.setObjectName("block_5_description")
        self.button_segregation = QtWidgets.QPushButton(self.block_5)
        self.button_segregation.setGeometry(QtCore.QRect(10, 75, 120, 35))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.button_segregation.setFont(font)
        self.button_segregation.setDefault(True)
        self.button_segregation.setObjectName("button_segregation")
        self.end_status = QtWidgets.QLabel(self.block_5)
        self.end_status.setGeometry(QtCore.QRect(130, 75, 500, 40))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        font.setStyleStrategy(QtGui.QFont.NoAntialias)
        self.end_status.setFont(font)
        self.end_status.setTextFormat(QtCore.Qt.AutoText)
        self.end_status.setAlignment(QtCore.Qt.AlignLeft)
        self.end_status.setWordWrap(True)
        self.end_status.setObjectName("end_status")
        self.end_button = QtWidgets.QPushButton(self.centralwidget)
        self.end_button.setGeometry(QtCore.QRect(510, 680, 120, 30))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.end_button.setFont(font)
        self.end_button.setLocale(QtCore.QLocale(QtCore.QLocale.Polish, QtCore.QLocale.Poland))
        self.end_button.setDefault(True)
        self.end_button.setFlat(False)
        self.end_button.setObjectName("end_button")
        self.setCentralWidget(self.centralwidget)

        # self.retranslateUi(MyWindow)
        QtCore.QMetaObject.connectSlotsByName(self)

        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "Aplikacja Segregująca pliki BOM'u"))
        self.block_1_title.setText(_translate("MainWindow", "Krok 1"))
        self.block_1_description.setText(_translate("MainWindow",
                                                    "Wybierz plik excel zawierający BOM i kliknij zweryfikuj aby sprawdzić czy jest on poprawny."))
        self.bom_button.setText(_translate("MainWindow", "Przeglądaj"))
        self.bom_verification_button.setText(_translate("MainWindow", "Zweryfikuj Poprawność"))
        self.bom_verification_info_text.setText(_translate("MainWindow", "..."))
        self.block_2_title.setText(_translate("MainWindow", "Krok 2"))
        self.block_2_description.setText(_translate("MainWindow",
                                                    "<html><head/><body><p align=\"justify\">Krok nie jest obowiązkowy. Kliknij przycisk aby program automatycznie policzył i wprowadził do arkusza Qty_Total.              Upewnij się, że plik nie jest obecnie otwarty.</p></body></html>"))
        self.qty_calculation_button.setText(_translate("MainWindow", "Policz Qty_Total"))
        self.qty_calculation_status.setText(_translate("MainWindow", "..."))
        self.block_3_title.setText(_translate("MainWindow", "Krok 3"))
        self.block_3_description.setText(_translate("MainWindow",
                                                    "Podaj ściezkę do miejsca gdzie bedą kopiowane posegregowane pliki i gdzie będzie zapisana lista brakujących plików."))
        self.button_destination_path.setText(_translate("MainWindow", "Przeglądaj"))
        self.block_4_title.setText(_translate("MainWindow", "Krok 4"))
        self.block_4_description.setText(_translate("MainWindow",
                                                    "Podaj ściezkę do lokalizacji, z której będą kopiowane pliki. Program przejrzy wszystkie podfoldery tego folderu."))
        self.button_source_path.setText(_translate("MainWindow", "Przeglądaj"))
        self.block_5_title.setText(_translate("MainWindow", "Krok 5"))
        self.block_5_description.setText(_translate("MainWindow",
                                                    "Kliknij przycisk Segreguj aby posegregować pliki zgodnie z podanymi wyżej danymi."))
        self.button_segregation.setText(_translate("MainWindow", "Segreguj"))
        self.end_status.setText(_translate("MainWindow", " "))
        self.end_button.setText(_translate("MainWindow", "Zakończ"))

        self.bom_button.clicked.connect(self.bom_button_clicked)
        self.bom_verification_button.clicked.connect(self.bom_verification_button_clicked)
        self.qty_calculation_button.clicked.connect(self.qty_calculation_button_clicked)
        self.button_destination_path.clicked.connect(self.button_destination_path_clicked)
        self.button_source_path.clicked.connect(self.button_source_path_clicked)
        self.button_segregation.clicked.connect(self.button_segregation_clicked)
        self.end_button.clicked.connect(self.end_button_clicked)

        self.bom_path = 'null'
        self.correct_bom_file_selected = False

    def bom_button_clicked(self):
        self.bom_path = QFileDialog.getOpenFileName(self, "Select File Name:", "D:\PROGRAMOWANIE",
                                                    " Excel files (*.xlsx *.xls)")
        print(self.bom_path)
        self.line_bom_path.setText(self.bom_path[0])
        font = QtGui.QFont()
        font.setItalic(False)
        self.line_bom_path.setFont(font)
        return self.bom_path

    def bom_verification_button_clicked(self):
        if self.bom_path == 'null':
            self.bom_verification_info_text.setText('NIE WYBRANO ŻADNEGO PLIKU ZAWIERAJĄCEGO BOM!!!')
        else:
            print(self.bom_path[0])
            self.path = self.bom_path[0]
            try:
                self.item_number, self.part_number, self.qty, self.qty_total, self.tch1, self.tch2, self.tch3, self.rysunek, self.max_row = excel_check(
                    self.path)
                self.bom_verification_info_text.setText('Wybrano poprawny plik BOM. Możesz przejść do kolejnego kroku.')
                self.correct_bom_file_selected = True
                return self.path, self.item_number, self.part_number, self.qty, self.qty_total, self.tch1, self.tch2, self.tch3, self.rysunek, self.max_row
            except:
                self.bom_verification_info_text.setText('Wybrany plik jest niepoprawny.')
                self.correct_bom_file_selected = False

    def qty_calculation_button_clicked(self):
        try:
            qty_total_calculation(self.path, self.item_number, self.qty, self.qty_total)
            self.qty_calculation_status.setText('Ilości w kolumnie Qty_Total zostały pomyślnie policzone.')
        except:
            self.qty_calculation_status.setText(
                "Nie udało się policzyć wartości. Sprawdź czy plik nie jest otwary w innym programie")

    def button_destination_path_clicked(self):
        self.destination_path = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.line_destination_path.setText(self.destination_path)

    def button_source_path_clicked(self):
        self.source_path = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.line_source_path.setText(self.source_path)

    def button_segregation_clicked(self):
        self.end_status.setText(" ")
        self.destination_path = self.line_destination_path.text()
        self.source_path = self.line_source_path.text()
        txt_file_name = "brakujace_pliki.txt"

        if self.correct_bom_file_selected:
            if os.path.exists(self.destination_path):
                if os.path.exists(self.source_path):

                    try:
                        self.end_status.setText("Czekaj...")
                        self.no_file_in_source, self.modification_time = temp_file_list(self.source_path,
                                                                                        self.destination_path,
                                                                                        self.path, self.part_number,
                                                                                        self.tch1, self.tch2, self.tch3,
                                                                                        self.rysunek, self.max_row)
                        self.end_status.setText("UKOŃCZONO TWORZENIE LISTY PLIKÓW DO KONWERSJI!!!")
                        print(self.modification_time)

                        # uruchomienie programu c_program w C# ktory konwertuje pliki solida
                        subprocess.call("solidworks_conversion_program.exe")
                        time.sleep(5)

                        # petla oczekujaca na nadpisanie pliku tymczasowego 'temp_file_txt.txt'
                        if self.modification_time != time.ctime(os.path.getmtime("temp_file_txt.txt")):
                            print('ROZPOCZYNAM PRZENOSZENIE PLIKÓW.')
                            move_files(self.destination_path, self.no_file_in_source)
                            os.remove("temp_file_txt.txt")
                        else:
                            self.end_status.setText("Wystapił problem w trakcie konwersji plików.")

                    except:
                        self.end_status.setText("Wystapił problem w trakcie tworzenie listy plików")

                    try:
                        txt_file_creation(self.destination_path, self.no_file_in_source, txt_file_name)
                        self.end_status.setText(
                            "UKOŃCZONO SEGREGACJE!!! Lista brakujących plików w pliku brakujące_pliki.txt")
                        txt_file_path = os.path.join(self.destination_path, txt_file_name)
                        webbrowser.open(txt_file_path)

                    except:
                        self.end_status.setText("Wystapił problem z zapisem brakujących plików do: " + txt_file_name)
                else:
                    self.end_status.setText("Nie wybrano ścieżki, z której mają być kopiowane pliki!")
            else:
                self.end_status.setText("Nie wybrano ścieżki, do której mają być kopiowane pliki!")
        else:
            self.end_status.setText("Nie wybrano BOM'u do segregacji lub wybrany plik nie został zweryfikowany.")

    def end_button_clicked(self):
        sys.exit(app.exec_())


def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())


window()
