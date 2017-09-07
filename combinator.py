#!/usr/bin/env python
#Skrypt składający w jeden zbiorczy plik inormacje z mediaplanów w bieżącym katalogu

import os
import re

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, get_column_interval
import pandas as pd
from pandas import DataFrame

def listMpFiles():
    """Tworzy listę plików z poszczególnymi mediaplanami w bieżącym katalogu"""
    mp_files = [f for f in os.listdir('.') if re.match(r'\b[a-zA-Z0-9].*.xls*', f)]
    num_files = len(mp_files)
    print('Znalazłem {} plików.'.format(num_files))
    return mp_files

def sheetToDataFrame(range_string, ws, kolumna):
    """Tworzy DataFrame z zakresu komórek 'A4':podana ostatnia kolumna i wiersz danych"""
    data_rows = []
    for row in ws[range_string]:
        data_rows.append([cell.value for cell in row])

    return DataFrame(data_rows, columns = get_column_interval('A', kolumna))

#if __name__ == __main__:

mediaplany = listMpFiles()
baza = Workbook()
arkuszBazy = baza.active
currentSheetDF = DataFrame()
finalDF = DataFrame()

arkusze = ('LIC - MP', 'LIC - H', 'SUM - MP', 'SUM - H', 'SP - MP', 'SP - H', 'MBA - MP', 'MBA - H',
           'Szkolenia - MP', 'Ogólne',)

#pętla ładująca poszczególne pliki z mediaplanami i sprawdzająca czy zawierają potrzebne arkusze
for filename in mediaplany:
    try:
        wb = load_workbook(filename, keep_vba=True)
    except:
        print('Błąd otwarcia pliku: {}'.format(filename))

    print('Sprawdzam plik: {}'.format(filename))

    for arkusz in arkusze: #ładuje poszczególne arkusze i zbiera z nich dane
        try:
            ws = wb[arkusz]
        except:
            print('\t! -> Nie znalazłem arkusza {} w pliku {}, pomijam.'.format(arkusz, filename))
            continue

        lastRow = 0

        #Ustala ostatni wiersz danych
        searchCol = ws['A']
        sumaInRow = 0
        for searchCell in searchCol:
            sumaInRow += 1
            if searchCell.value == 'SUMA':
                lastRow = sumaInRow - 1
                print('\tOstatni wiersz danych w {} to {}'.format(arkusz, lastRow))
            else: continue

        #Ustala ostatnią kolumnę danych
        searchRow = ws[3]
        sumaInCol = 0
        for searchCell in searchRow:
            sumaInCol += 1
            if searchCell.value != 'ga:source':
                continue
            else:
                kolumna = get_column_letter(sumaInCol)
                print('\tOstatnia kolumna danych to: {}'.format(kolumna))

        final_data_cell = str(kolumna) + str(lastRow)
        cellsRange = 'A4' + ':' + str(final_data_cell)
        currentSheetDF = sheetToDataFrame(cellsRange, ws, kolumna)
        finalDF = finalDF.append(currentSheetDF, ignore_index=True)

    print('Plik: {} OK!'.format(filename))

writer = pd.ExcelWriter('baza.xlsx', engine='xlsxwriter')

finalDF.to_excel(writer,
                 sheet_name='Baza',
                 index=False)

workbook = writer.book
worksheet = writer.sheets['Baza']
writer.save()

