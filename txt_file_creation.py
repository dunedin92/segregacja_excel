#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

import os


def txt_file_creation(destination, file_list, txt_file_name):

    txt_file_path = os.path.join(destination, txt_file_name)
    plik = open(txt_file_path, "w", encoding='utf8')

    for name in file_list:
        plik.write(name)
        plik.write("\n")
    plik.close()

    return 0
