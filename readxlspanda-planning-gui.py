# doc: https://pythonhosted.org/xlrd3/


import codecs
import locale
import os
from functools import partial
from tkinter import *

import pandas as pd
import win32com.client
from pandastable import Table, config
from pretty_html_table import build_table
from tkcalendar import Calendar

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
df = '%a %d/%m/%Y'
# locale.setlocale(locale.LC_TIME, 'fr_FR')
# locale.setlocale(locale.LC_ALL, 'fr')


intervenant = 'Pilote'

# promo starts with
recipient = 'phofmannh@gmail.com'
signature_name = 'new'

# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# promos = ('ASR1 23-24', 'ASR2-FFASR 23-24', 'MAALSI 23-25', 'ASR2-FFASR 24-25')

# plannings = (
#     #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\14 - ASR 22-23 - TL2XN204\02 - Rythme et Planning\Planning ASR01- 2022-2023.xlsx",
#     #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\15 - ASR_FFASR 22-23 - TL2XN214\02 - Rythme et Planning\Planning ASR02-FFASR - 2022-2023.xlsx",
#     r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\16 - ASR1 23-24 - TL3XN204\02 - Rythme et Planning\Planning ASR1- 2023-2024.xlsx",
#     r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\17 - ASR2-FFASR2 23-24 -TL3XN206\02 - Rythme et Planning\Planning ASR2-FFASR - 2023-2024.xlsx",
#     r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\6 - MAALSI\02 - MAALSI 23-25\02 - Rythme et Planning\PLANNING MAALSI 2023-2025.xlsx",
#     r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\18 - ASR1 24-25\02 - Rythme et Planning\Planning ASR1 - 2024-2025.xlsx",
#     r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\19 - ASR2-FFASR2 24-25\02 - Rythme et Planning\Planning ASR2-FFASR - 2024-2025.xlsx"
#     #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\3 - CDA\14 - CDA 23-24 - TL3XN201\02 - Rythme et Planning\PLANNING CDA 2023-2024.xlsm"
# )
plannings = {
    'ASR1 23-24': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\16 - ASR1 23-24 - TL3XN204\02 - Rythme et Planning\Planning ASR1- 2023-2024.xlsx",
    'ASR2 23-24': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\17 - ASR2-FFASR2 23-24 -TL3XN206\02 - Rythme et Planning\Planning ASR2-FFASR - 2023-2024.xlsx",
    'MAALSI 23-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\6 - MAALSI\02 - MAALSI 23-25\02 - Rythme et Planning\PLANNING MAALSI 2023-2025.xlsx",
    'ASR1 24-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\18 - ASR1 24-25\02 - Rythme et Planning\Planning ASR1 - 2024-2025.xlsx",
    'ASR2 24-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\19 - ASR2-FFASR2 24-25\02 - Rythme et Planning\Planning ASR2-FFASR - 2024-2025.xlsx"
    #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\3 - CDA\14 - CDA 23-24 - TL3XN201\02 - Rythme et Planning\PLANNING CDA 2023-2024.xlsm"
}

df_plannings = {}

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
df_all_plannings = pd.DataFrame()  # all the plannings
df_filtered_plannings = pd.DataFrame()  # filtered planning

signature_code = str()


# def load_all_plannings():
#     global df_all_plannings
#     for planning in plannings:
#         print("loading " + planning)
#         df = pd.read_excel(planning, skiprows=0, header=1)
#         # print(df[['Promo', 'AM-PM','Date', 'Charge',  'Intervenants' ,'Réf.', 'Module']])
#         df_all_plannings = df_all_plannings._append(
#             df[['Promo', 'AM-PM', 'Date', 'Charge', 'Intervenants', 'Réf.', 'Module', 'Confirmation']])
#         del df
#     # replace empty cells by empty string (not sure why....)
#     # df_all_plannings.fillna('', inplace=True)
#     # put a readable date as column jour
#
#     df_all_plannings.fillna('', inplace=True)
#
#     df_all_plannings = df_all_plannings[df_all_plannings['Réf.'] != '']
#
#     #    print(df_all_plannings[df_all_plannings['Intervenants'] == 'FOLIO Brice'])
#     df_all_plannings['Jour'] = df_all_plannings['Date'].dt.strftime('%a %d/%m/%Y')
#
#     df_all_plannings.sort_values('Date')
#     df_all_plannings.reset_index(inplace=True, drop=True)
#
#     return

def load_all_plannings():
    global df_plannings
    for p in plannings.keys():
        print("loading " + p + ": " + plannings[p])
        df1 = pd.read_excel(plannings[p], skiprows=0, header=1)
        # print(df[['Promo', 'AM-PM','Date', 'Charge',  'Intervenants' ,'Réf.', 'Module']])
        df2 = df1[['Promo', 'AM-PM', 'Date', 'Charge', 'Intervenants', 'Réf.', 'Module', 'Confirmation']]
        del df1
        df2.fillna('', inplace=True)
        df2 = df2[df2['Réf.'] != '']
        #    print(df_all_plannings[df_all_plannings['Intervenants'] == 'FOLIO Brice'])
        df2['Jour'] = df2['Date'].dt.strftime('%a %d/%m/%Y')
        df2.sort_values('Date')
        df2.reset_index(inplace=True, drop=True)
        df_plannings[p] = df2
    return


# def get_promos(plannings):
#     p = ()
#
#     p = df_all_plannings['Promo'].tolist()
#     # p = list(dict.fromkeys(p))  # remove duplicates
#     p = list(set(p))
#     for x in p:
#         if (x == ''):
#             print(df_all_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']]
#                   .to_string(index=False, justify='left'))
#     print(p)
#     return p


def get_intervenants(plannings):
    l = df_all_plannings['Intervenants'].tolist()
    l = list(dict.fromkeys(l))  # remove duplicates
    l.sort()
    print(l)
    return l


def filter_plannings(frame):
    global df_filtered_plannings
    result_df = pd.DataFrame()

    # delete all rows
    df_filtered_plannings = df_filtered_plannings.head(0)

    print("date ", filter_param['date'])
    print("promos", filter_param['promos'])
    print('intervenant', filter_param['intervenant'])
    for p, df in df_plannings.items():
        if filter_param['promos'][p] == True:
            df_filtered_plannings = df_filtered_plannings._append(df.loc[
                                                                      (df['Intervenants'].str.contains(
                                                                          filter_param['intervenant'])) &
                                                                      # (df_all_plannings['Promo'].str.contains(dict['promo'])) &
                                                                      (df['Date'] > pd.to_datetime(filter_param['date'],
                                                                                                   dayfirst=True))
                                                                      ])
    df_filtered_plannings.sort_values('Date')
    df_filtered_plannings.reset_index(inplace=True, drop=True)
    pt = Table(frame,
               dataframe=df_filtered_plannings[
                   ['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module', 'Confirmation']],
               showtoolbar=True, showstatusbar=True)
    # set some options
    options = {'colheadercolor': 'blue', 'floatprecision': 1, 'fontsize': 8, 'cellwidth': 40}
    config.apply_options(options, pt)
    pt.show()

    # coloring cells
    # mask_1 = pt.model.df['Confirmation'] != ''
    # pt.setColorByMask('Confirmation', mask_1, 'red')
    df_filtered_plannings.reset_index(inplace=True, drop=True)

    r = ()
    r = df_filtered_plannings.index[df_filtered_plannings['Confirmation'] == ''].tolist()
    r.sort()
    # print(r)
    # print(df_filtered_plannings[df_filtered_plannings['Confirmation'] == ''])
    pt.setRowColors(rows=r, clr='#BBBBBB', cols='all')

    return


# load signature
def load_signature(signature_name):
    global signature_code

    sig_files_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '_fichiers\\'
    sig_html_path = 'AppData\Roaming\Microsoft\Signatures\\' + signature_name + '.htm'
    signature_path = os.path.join((os.environ['USERPROFILE']),
                                  sig_files_path)  # Finds the path to Outlook signature files with signature name "Work"
    html_doc = os.path.join((os.environ['USERPROFILE']),
                            sig_html_path)  # Specifies the name of the HTML version of the stored signature
    html_doc = html_doc.replace('\\\\', '\\')
    html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore')
    signature_code = html_file.read()  # Writes contents of HTML signature file to a string
    signature_code = signature_code.replace((signature_name + '_fichiers/'),
                                            signature_path)  # Replaces local directory with full directory path
    html_file.close()


# -----------------------------------------------------------------------------
def print_planning(p):
    # print tabulated on stdout
    with pd.option_context('display.colheader_justify', 'left', 'display.max_columns', None, 'display.width', None,
                           'display.float_format', "{:.2f}  ".format):
        print(
            p[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module', 'Confirmation']].to_string(
                index=False,
                justify='left'))


def display_plannings():
    ### https://pypi.org/project/pretty-html-table/
    html_beautiful_table = build_table(
        df_filtered_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
        'blue_light',
        font_family='Open Sans , sans-serif',
        font_size='12px',
        width_dict=['120px', '30px', '20px', '180px', '140px', '60px', '400px'])

    # write a copy in C:\Users\phofmann\Downloads\planning2.html
    with open(r"C:\Users\phofmann\Downloads\planning2.html", 'w') as f:
        f.write(html_beautiful_table)


def prepare_email():
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0  # size of the new email
    load_signature(signature_name)

    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'CESI: Confirmation planning interventions'
    newmail.To = recipient
    # newmail.CC='xyz@gmail.com'

    newmail.BodyFormat = 3  # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    ### https://pypi.org/project/pretty-html-table/
    html_beautiful_table = build_table(
        df_filtered_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
        'blue_light',
        font_family='Open Sans , sans-serif',
        font_size='12px',
        width_dict=['120px', '30px', '20px', '180px', '140px', '60px', '400px'])
    newmail.HTMLBody = html_mail + html_beautiful_table + " <br><br><br> " + signature_code
    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)
    # To display the mail before sending it
    newmail.Display()
    return


################################
def call_filter_planning(frame, cal, ie):
    date = date = cal.get_date()
    if date == '':
        date = today
    filter_param['date'] = date
    filter_param['intervenant'] = ie.get()
    # set all promo to false
    for p in filter_param['promos'].keys():
        filter_param['promos'][p] = False
    for i in promoListbox.curselection():
        print(promoListbox.get(i))
        filter_param['promos'][promoListbox.get(i)] = True

    filter_plannings(frame)
    # display_plannings()
    return


def get_start_date(*args, frame, d, p, i):  # triggered on Button Click
    #   print("date: {}".format(d.get()))
    #   print("promo: {}".format(p.get()))
    #   print("intervenant: {}".format(i.get()))
    date = cal.get_date()  # read and display date
    print(date)
    date = (d.get())
    if date == '':
        date = today
    filter_param['date'] = date
    filter_param['intervenant'] = i.get()
    # set all promo to false
    for p in filter_param['promos'].keys():
        filter_param['promos'][p] = False
    for i in promoListbox.curselection():
        print(promoListbox.get(i))
        filter_param['promos'][promoListbox.get(i)] = True
    filter_plannings(frame)
    # label_result.config(text="Result = %d" % result)
    # display_plannings()
    # del pt
    return


################# main
filter_param = {
    'date': Calendar.date.today().strftime("%d/%m/%y"),
    'promos': {},
    'intervenant': ''
}

top = Tk()
top.title("gestion plannings")
top.geometry("1200x700")
f1 = Frame(top, height=200)  # for the buttons and the entries
f1.pack(fill='both', side=TOP)
f2 = Frame(top)  # for the tables
f2.pack(fill='both', side=BOTTOM)

# promoLabel = Label(f1, text="Promo contient")
# promoLabel.place(x=1, y=40)
intervenantLabel = Label(f1, text="Intervenant")
intervenantLabel.place(x=1, y=70)

# promoVar = StringVar()
intervenantVar = StringVar()

# promoEntry = Entry(f1, textvariable=promoVar)
# promoEntry.place(x=100, y=40)
intervenantEntry = Entry(f1, textvariable=intervenantVar)
intervenantEntry.place(x=100, y=70)

start_date = StringVar(f1, Calendar.date.today().strftime("%d/%m/%y"))

# print(today.year, today.month, today.day)
# print(today)
cal = Calendar(f1,
               selectmode='day',
               date_pattern='dd/mm/yyyy',
               year=today.year, month=today.month, day=today.day,
               locale='fr',
               # start_date=Calendar.date.today().strftime("%d/%m/%y"),
               textvariable=start_date)
cal.place(x=400, y=10)
start_date.trace('w', lambda *_, f=f2, d=start_date, i=intervenantVar: get_start_date(*_, frame=f, d=d, p=p,
                                                                                      i=i))

f1.pack(fill='both', side=TOP)

print("start loading plannings")
load_all_plannings()
print("plannings loaded")
# print_planning(df_all_plannings)


# promoEntry.bind("<Return>", lambda e: call_filter_planning(f2, cal, promoEntry, intervenantEntry))
intervenantEntry.bind("<Return>", lambda e: call_filter_planning(f2, cal, intervenantEntry))

promoListbox = Listbox(f1, height=len(plannings), width=20, selectmode=MULTIPLE)
promoListbox.place(x=250, y=10)

i = 0
for p in plannings.keys():
    promoListbox.insert(i + 1, p)
    filter_param['promos'][p] = True;
    i = i + 1

f1.pack(fill='both', side=TOP)

c1 = partial(call_filter_planning, f2, cal, intervenantEntry)
filterButton = Button(f1, text="filter plannings", activebackground='green', command=c1)
filterButton.place(x=50, y=100)

c2 = partial(prepare_email)
prepareEmailButton = Button(f1, text="prepare email", activebackground='blue', command=c2)
prepareEmailButton.place(x=150, y=100)

# pt = Table(f2, dataframe=df_all_plannings[['Jour', 'AM-PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
#           showtoolbar=True, showstatusbar=True)
# set some options
# options = {'colheadercolor': 'blue', 'floatprecision': 1, 'fontsize': 8, 'cellwidth': 40}
# config.apply_options(options, pt)
# pt.show()
filter_plannings(f2)

top.mainloop()
