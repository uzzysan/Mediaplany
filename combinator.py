#!/usr/bin/env python
#Skrypt składający w jeden zbiorczy plik inormacje z mediaplanów w bieżącym katalogu

import os, re
from openpyxl import load_workbook, Workbook
import numpy as np
import pandas as pd
from pandas import Series, DataFrame

def listMpFiles():
    "Tworzy listę plików z poszczególnymi mediaplanami w bieżącym katalogu"
    mpFiles = [f for f in os.listdir('.') if re.match(r'\b[a-zA-Z0-9].*.xls*', f)]
    numFiles = len(mpFiles)
    print('Znalazłem {} plików.'.format(numFiles))
    return mpFiles

#if __name__ == __main__:

mediaplany = listMpFiles()

arkusze = ('LIC - MP', 'LIC - H', 'SUM - MP', 'SUM - H', 'SP - MP', 'SP - H', 'MBA - MP', 'MBA - H', 'Szkolenia - MP', 'Ogólne',)

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
        
        searchCol = ws['A']
        sumaInRow = 0
        for searchCell in searchCol:
            sumaInRow += 1
            if searchCell.value == 'SUMA':
                lastRow = sumaInRow - 1
                print('\tOstatni wiersz danych w {} to {}'.format(arkusz, lastRow))
            else: continue

        wiersz = 0
        for row in ws.iter_rows(min_row=4, max_col=113, max_row=lastRow):
            wiersz += 1
        print('\tArkusz {}, wczytano {} wierszy danych.'.format(arkusz, wiersz))

    print('Plik: {} OK!'.format( filename))

