# read "C:\Users\phofmann\Dropbox\private\finances\aok2_2022-11-27.ods"

import pandas as pd
import datetime
from pandas_ods_reader import read_ods
import re
from numpy import *
import operator

#odsfile = r"C:\Users\phofmann\Dropbox\private\finances\aok2_2022-11-27.ods"
odsfile = r"D:\data\Dropbox\Dropbox\private\finances\aok2_2022-11-27.ods"

# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.date_dayfirst', True)
pd.set_option('display.max_colwidth', 80)

df = pd.DataFrame()

# Read the ODS file into a DataFrame
df = read_ods(odsfile, "CM")
df["date"] = pd.to_datetime(df['Date de valeur'])
df["date_str"] = df['date'].dt.strftime('%d/%m/%Y')
df.fillna('', inplace=True)

# Print the DataFrame
# .strftime('%d/%m/%Y')


# with pd.option_context('display.colheader_justify', 'right', 'display.max_rows', 100, 'display.max_columns', None, 'display.width', None,'display.float_format', "{:.2f}  ".format):
# print(df[["date_str", "Quoi", "description", "montan €"]])

# https://www.linkedin.com/pulse/simple-accounting-automation-using-naive-bayes-ryan-park-cpa-cga/

libelle = []
categories = []
libelle = df["description"].tolist()
# print (libelle)
categories = df["Quoi"].tolist()
# print(categories)

# cree la liste du vocabulaire dans les libellés
vocabulaire = set([])
for item in libelle:
    tokens = re.split(r'\W*', item)
    # print (tokens)
    tokens2 = [tok.lower() for tok in tokens if len(tok) > 1]
    vocabulaire = vocabulaire | set(tokens2)
# print (vocabulaire)

libelle_adj = []
for item in libelle:
    tokens = re.split(r'\W*', item)
    print(tokens)
    tokens2 = [tok.lower() for tok in tokens if len(tok) > 1]
    libelle_adj.append(tokens2)
    print(tokens2)

print(libelle_adj)
