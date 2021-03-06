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

def clean_df(dataframes):
    summary_text = 'SUMA'
    for df in dataframes:
        index_after_suma = df.index.str.startswith(summary_text).cumsum()
        yield df.loc[~index_after_suma, :]

mediaplany = listMpFiles()
arkusze = ('LIC - MP', 'LIC - H', 'SUM - MP', 'SUM - H', 'SP - MP', 'SP - H', 'MBA - MP', 'MBA - H', 'Szkolenia - MP', 'Ogólne')
finalDF = pd.DataFrame()

for fileName in mediaplany:
    print('Otwieram plik: {}...'.format(fileName))
    for arkusz in arkusze:
        try:
            df = pd.read_excel(fileName,
                               sheetname=arkusz,
                               header=None,
                               skiprows=3,
                               index_col=None,
                               skip_footer=0,
                               parse_cols='A:J,AB:CC,CE:DJ',
                               na_values='')
            clean_df(df)
            df.dropna(axis=0, how='all')
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

