#!/usr/bin/env python3
# coding: utf-8
# -*- coding: utf-8 -*-

import openpyxl


def qty_total_calculation(bom_path, kolumna_item_number, kolumna_qty, kolumna_qty_total):
    # otworzenie arkusza BOM'u
    wb = openpyxl.load_workbook(bom_path)
    type(wb)
    arkusze = wb.sheetnames
    sheet = wb[arkusze[0]]
    max_row = sheet.max_row

    print(kolumna_item_number)
    print(kolumna_qty)
    print(kolumna_qty_total)

    # liczenie ilosci kropek i przecinkow w numerze
    for i in range(2, max_row + 1):
        cell_value_actual = sheet.cell(row=i, column=kolumna_item_number).value
        cell_value_last = sheet.cell(row=i - 1, column=kolumna_item_number).value
        print(cell_value_actual)
        cell_value_actual = str(cell_value_actual)
        cell_value_last = str(cell_value_last)
        qt_dots_actual = (cell_value_actual.count(".")) + (cell_value_actual.count(","))
        qt_dots_last = (cell_value_last.count(".")) + (cell_value_last.count(","))
        # print(qt_dots_actual)
        # print('roznica w zagniezdzeniu:')
        # print(qt_dots_actual - qt_dots_last)

        if qt_dots_actual == 0:
            qty_i = sheet.cell(row=i, column=kolumna_qty).value
            sheet.cell(row=i, column=kolumna_qty_total).value = qty_i
            # print('wpisujemy wartosc qty_total:')
            # print(sheet.cell(row=i, column=kolumna_qty_total).value)
        elif qt_dots_actual - qt_dots_last == 1:
            qty_i_last = sheet.cell(row=i - 1, column=kolumna_qty_total).value
            qty_i = sheet.cell(row=i, column=kolumna_qty).value
            sheet.cell(row=i, column=kolumna_qty_total).value = qty_i_last * qty_i
            # print('wpisujemy wartosc qty_total:')
            # print(sheet.cell(row=i, column=kolumna_qty_total).value)
        elif qt_dots_actual - qt_dots_last == 0 or qt_dots_actual - qt_dots_last <= -1:
            # print('ten sam poziom zagniezdzenia co wyzej')
            for j in range(i - 1, 0, -1):
                temporary_cell_value = sheet.cell(row=j, column=kolumna_item_number).value
                temporary_cell_value = str(temporary_cell_value)
                qt_dots_temporary = (temporary_cell_value.count(".")) + (temporary_cell_value.count(","))
                if qt_dots_actual - qt_dots_temporary == 1:
                    # print('znaleziono nizszy poziom zagniezdzenia')
                    qty_i = sheet.cell(row=i, column=kolumna_qty).value
                    qty_i_temporary = sheet.cell(row=j, column=kolumna_qty_total).value
                    sheet.cell(row=i, column=kolumna_qty_total).value = qty_i * qty_i_temporary
                    # print('wpisujemy wartosc qty_total:')
                    # print(sheet.cell(row=i, column=kolumna_qty_total).value)
                    break

    wb.save(bom_path)
    return True
