# doc: https://pythonhosted.org/xlrd3/


import locale

import pandas as pd
from pretty_html_table import build_table
import win32com.client
import os
import codecs

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

intervenant = 'FOLIO'
promo = 'ASR'
recipient = 'phofmannh@gmail.com'
signature_name = 'new'



# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

plannings = (
 #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\14 - ASR 22-23 - TL2XN204\02 - Rythme et Planning\Planning ASR01- 2022-2023.xlsx",
 #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\15 - ASR_FFASR 22-23 - TL2XN214\02 - Rythme et Planning\Planning ASR02-FFASR - 2022-2023.xlsx",
    r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\16 - ASR1 23-24 - TL3XN204\02 - Rythme et Planning\Planning ASR1- 2023-2024.xlsx",
    r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\17 - ASR2-FFASR2 23-24 -TL3XN206\02 - Rythme et Planning\Planning ASR2-FFASR - 2023-2024.xlsx",
    r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\6 - MAALSI\02 - MAALSI 23-25\02 - Rythme et Planning\PLANNING MAALSI 2023-2025.xlsx",
    r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\19 - ASR2-FFASR2 24-25\02 - Rythme et Planning\Planning ASR2-FFASR - 2024-2025.xlsx",
    r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\18 - ASR1 24-25\02 - Rythme et Planning\Planning ASR1 - 2024-2025.xlsx"
)
# Open the Workbook

html_mail = """\
<html>
<body>
<p>Bonjour <br>
   Voici, pour confirmation, le planning de vos interventions prévues ensemble<br>
   Merci de me retourner un email pour validation<br>
   
   Bien cordialement<br>
   Philippe<br><br>
</p>
</body>
</html>
"""

today = pd.Timestamp.today()
df_plannings = pd.DataFrame()
signature_code = str()

def load_plannings( _intervenant,_promo,_date):
    global df_plannings

    df_all = pd.DataFrame()
    for planning in plannings:
        df = pd.read_excel(planning)
        #df.fillna('', inplace=True)
        # print(df[['Promo', 'AM-PM','Date', 'Charge',  'Intervenants' ,'Réf.', 'Module']])
        df_all = df_all._append(df[['Promo', 'AM-PM', 'Date', 'Charge', 'Intervenants', 'Réf.', 'Module']])
        del df

    df_all.fillna('', inplace=True)
    #df_all.sort_values('Date')
    df_plannings = df_all.loc[
        (df_all['Intervenants'].str.startswith(_intervenant)) &
        (df_all['Promo'].str.startswith(_promo)) &
        (df_all['Date'] > _date)
        ].sort_values('Date')
    df_plannings['Jour'] = df_plannings['Date'].dt.strftime('%a %d/%m/%Y')


#load signature
def load_signature(signature_name):
    global signature_code

    sig_files_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '_fichiers\\'
    sig_html_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '.htm'

    signature_path = os.path.join((os.environ['USERPROFILE']), sig_files_path) # Finds the path to Outlook signature files with signature name "Work"
    html_doc = os.path.join((os.environ['USERPROFILE']),sig_html_path)     #Specifies the name of the HTML version of the stored signature
    html_doc = html_doc.replace('\\\\', '\\')

    html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
    signature_code = html_file.read()               #Writes contents of HTML signature file to a string

    signature_code = signature_code.replace((signature_name + '_fichiers/'), signature_path)      #Replaces local directory with full directory path
    html_file.close()

#-----------------------------------------------------------------------------

#load plannings
load_plannings( intervenant,promo,today)

# print tabulated on stdout
with pd.option_context('display.colheader_justify', 'left', 'display.max_columns', None, 'display.width', None,'display.float_format', "{:.2f}  ".format):
    print(df_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']].to_string(index=False,justify='left'))

### https://pypi.org/project/pretty-html-table/
html_beautiful_table=build_table(df_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
                                 'blue_light',
                                font_family='Open Sans , sans-serif',
                                 font_size='12px',
                                 width_dict=['120px','30px', '20px', '180px','140px', '60px', '400px'])

#write a copy in C:\Users\phofmann\Downloads\planning2.html
with open(r"C:\Users\phofmann\Downloads\planning2.html", 'w') as f:
    f.write(html_beautiful_table)

ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0  # size of the new email
load_signature(signature_name)

newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'CESI: Confirmation planning interventions'
newmail.To = recipient
# newmail.CC='xyz@gmail.com'

newmail.BodyFormat = 3  # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newmail.HTMLBody = html_mail  + html_beautiful_table + signature_code
# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)
# To display the mail before sending it
newmail.Display()