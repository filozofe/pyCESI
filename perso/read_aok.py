#read "C:\Users\phofmann\Dropbox\private\finances\aok2_2022-11-27.ods"

import pandas as pd
import datetime
from pandas_ods_reader import read_ods

odsfile = r"C:\Users\phofmann\Dropbox\private\finances\aok2_2022-11-27.ods"

# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.date_dayfirst', True)
pd.set_option('display.max_colwidth', 80)

df = pd.DataFrame()

# Read the ODS file into a DataFrame
df = read_ods(odsfile,"CM")
df["date"]=pd.to_datetime(df['Date de valeur'])
df["date_str"]=df['date'].dt.strftime('%d/%m/%Y')

# Print the DataFrame
#.strftime('%d/%m/%Y')


#with pd.option_context('display.colheader_justify', 'right', 'display.max_rows', 100, 'display.max_columns', None, 'display.width', None,'display.float_format', "{:.2f}  ".format):
print(df [["date_str","Quoi","description","montan €"]])
