#!/usr/bin/env python
#Skrypt składający w jeden zbiorczy plik inormacje z mediaplanów w bieżącym katalogu

import os, re, datetime
import pandas as pd
import numpy as np
from pandas import Series, DataFrame

def listMpFiles():
    "Tworzy listę plików z poszczególnymi mediaplanami w bieżącym katalogu"
    mpFiles = [f for f in os.listdir('.') if re.match(r'\b[a-zA-Z0-9].*.xls*', f)]
    numFiles = len(mpFiles)
    print('Znalazłem {} plików excela.'.format(numFiles))
    return mpFiles

mediaplany = listMpFiles()
arkusze = ['LIC - MP', 'LIC - H', 'SUM - MP', 'SUM - H', 'SP - MP', 'SP - H', 'MBA - MP', 'MBA - H', 'Szkolenia - MP', 'Ogólne']

finalDF = pd.DataFrame()

for fileName in mediaplany:
    print('Otwieram plik: {}...'.format(fileName))
    for arkusz in arkusze:
        try:
            df = pd.read_excel(fileName,
                               sheetname=arkusz,
                               header=0,
                               skiprows=2,
                               skip_footer=0,
                               na_values='')

            finalDF = finalDF.append(df)
            print('\t+ Arkusz {} OK!'.format(arkusz))
        except:
            print('\t! Nie znalazłem arkusza {} w pliku {}, pomijam.'.format(arkusz, fileName))
            continue
        continue
    continue

writer = pd.ExcelWriter('baza.xlsx', engine='xlsxwriter')

finalDF.to_excel(writer,
                 sheet_name='Baza',
                 index=False)

workbook  = writer.book
worksheet = writer.sheets['Baza']
writer.save()

