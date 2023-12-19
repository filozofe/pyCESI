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
recipient = ('')
signature_name = 'new'

# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# promos = ('ASR1 23-24', 'ASR2-FFASR 23-24', 'MAALSI 23-25', 'ASR2-FFASR 24-25')

planning_file = r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\plannings_informatique.xlsx"

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
# plannings = {
#     'ASR1 23-24': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\16 - ASR1 23-24 - TL3XN204\02 - Rythme et Planning\Planning ASR1- 2023-2024.xlsx",
#     'ASR2 23-24': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\17 - ASR2-FFASR2 23-24 -TL3XN206\02 - Rythme et Planning\Planning ASR2-FFASR - 2023-2024.xlsx",
#     'MAALSI 23-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\6 - MAALSI\02 - MAALSI 23-25\02 - Rythme et Planning\PLANNING MAALSI 2023-2025.xlsx",
#     'ASR1 24-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\18 - ASR1 24-25\02 - Rythme et Planning\Planning ASR1 - 2024-2025.xlsx",
#     'ASR2 24-25': r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\2 - RISR_ASR\19 - ASR2-FFASR2 24-25\02 - Rythme et Planning\Planning ASR2-FFASR - 2024-2025.xlsx"
#     #   r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\3 - CDA\14 - CDA 23-24 - TL3XN201\02 - Rythme et Planning\PLANNING CDA 2023-2024.xlsm"
# }

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


def load_all_plannings():
    global df_plannings

    # lire fichier xls contenant le nom des promos et le path des fichiers plannings
    plannings = pd.read_excel(planning_file, skiprows=0, header=0)
    for i in range(0, len(plannings)):
        # print("loading " + plannings.iloc[i]['promo'] + ": " + plannings.iloc[i]['fichier'])
        print("loading " + plannings.iloc[i]['promo'])
        if plannings.iloc[i]['HeaderLine'] == 1:
            hl = 1
        else:
            hl = 0
        try:
            df1 = pd.read_excel(plannings.iloc[i]['fichier'], skiprows=0, header=hl)
        except FileNotFoundError:
            msg = f"can't find file: {plannings.iloc[i]['fichier']}"
            print(msg)
        else:
            # print(df[['Promo', 'AM-PM','Date', 'Charge',  'Intervenants' ,'Réf.', 'Module']])
            if 'Confirmation' in df1:
                # la colonne 'confirmation' est présente
                df2 = df1[['AM-PM', 'Date', 'Charge', 'Intervenants', 'Réf.', 'Module', 'Confirmation']]
            else:
                # sinon, creer une colonne vide
                df2 = df1[['AM-PM', 'Date', 'Charge', 'Intervenants', 'Réf.', 'Module']]
                df2 = df2.assign(Confirmation='')
            del df1
            df2['Promo'] = plannings.iloc[i]['promo']
            df2.fillna('', inplace=True)
            df2 = df2[df2['Réf.'] != '']
            #    print(df_all_plannings[df_all_plannings['Intervenants'] == 'FOLIO Brice'])
            df2['Jour'] = df2['Date'].dt.strftime('%a %d/%m/%Y')
            df2.sort_values('Date')
            df2.reset_index(inplace=True, drop=True)
            df_plannings[plannings.iloc[i]['promo']] = df2
    return


def get_intervenants():
    l = []
    for df in df_plannings.values():
        l.extend(df['Intervenants'].tolist())
    # l = list(dict.fromkeys(l))  # remove duplicates
    l = list(set(l))
    l.sort()
    # print(l)
    return l


def filter_plannings(frame):
    global df_filtered_plannings
    result_df = pd.DataFrame()

    # delete all rows
    df_filtered_plannings = df_filtered_plannings.head(0)

    # print("date ", filter_param['date'])
    # print("promos", filter_param['promos'])
    # print('intervenant', filter_param['intervenant'])
    for p, df in df_plannings.items():
        if filter_param['promos'][p] == True:
            df_filtered_plannings = df_filtered_plannings._append(df.loc[
                                                                      (df['Intervenants'].str.contains(
                                                                          filter_param['intervenant'])) &
                                                                      # (df_all_plannings['Promo'].str.contains(dict['promo'])) &
                                                                      (df['Date'] > pd.to_datetime(filter_param['date'],
                                                                                                   dayfirst=True))
                                                                      ])
    df_filtered_plannings.sort_values('Date', inplace=True)
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
    pt.setRowColors(rows=r, clr='#FFFFE0', cols='all')

    return


def find_collision(frame):
    global df_filtered_plannings
    result_df = pd.DataFrame()

    # delete all rows
    df_filtered_plannings = df_filtered_plannings.head(0)

    # print("date ", filter_param['date'])
    # print("promos", filter_param['promos'])
    # print('intervenant', filter_param['intervenant'])
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

    df_filtered_plannings['collision'] = df_filtered_plannings.duplicated(subset=['Jour', 'Intervenants', 'AM-PM'],
                                                                          keep=False)
    df_filtered_plannings = df_filtered_plannings[df_filtered_plannings.collision == True]
    # remove from duplicate the following 'Intervenants'
    exclude = ['', '?', 'Pilote', 'pilote', 'Autonomie']
    df_filtered_plannings = df_filtered_plannings[~df_filtered_plannings.Intervenants.isin(exclude)]
    df_filtered_plannings.sort_values('Date', inplace=True)
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
    pt.setRowColors(rows=r, clr='#FFFFE0', cols='all')
    statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " collisions")

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
        # print(promoListbox.get(i))
        filter_param['promos'][promoListbox.get(i)] = True

    filter_plannings(frame)
    statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " lignes selectionnées")
    # display_plannings()
    return


def get_start_date(*args, frame, d, p, i):  # triggered on Button Click
    #   print("date: {}".format(d.get()))
    #   print("promo: {}".format(p.get()))
    #   print("intervenant: {}".format(i.get()))
    date = cal.get_date()  # read and display date
    # print(date)
    date = (d.get())
    if date == '':
        date = today
    filter_param['date'] = date
    filter_param['intervenant'] = i.get()
    # set all promo to false
    for p in filter_param['promos'].keys():
        filter_param['promos'][p] = False
    for i in promoListbox.curselection():
        # print(promoListbox.get(i))
        filter_param['promos'][promoListbox.get(i)] = True
    filter_plannings(frame)
    statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " lignes selectionnées")
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
top.geometry("1200x800")
f1 = Frame(top, height=250)  # for the buttons and the entries
f1.pack(fill='both', side=TOP)
f2 = Frame(top)  # for the tables
f2.pack(fill='both', side=BOTTOM)

# promoLabel = Label(f1, text="Promo contient")
# promoLabel.place(x=1, y=40)
intervenantLabel = Label(f1, text="Intervenant")
intervenantLabel.place(x=1, y=10)
intervenantVar = StringVar()
intervenantEntry = Entry(f1, textvariable=intervenantVar)
intervenantEntry.place(x=80, y=10)

statusLabel = Label(f1, text="starting")
statusLabel.place(x=1, y=190)

start_date = StringVar(f1, Calendar.date.today().strftime("%d/%m/%y"))

calendarLabel = Label(f1, text="Date début")
calendarLabel.place(x=400, y=0)
cal = Calendar(f1,
               selectmode='day',
               date_pattern='dd/mm/yyyy',
               year=today.year, month=today.month, day=today.day,
               locale='fr',
               # start_date=Calendar.date.today().strftime("%d/%m/%y"),
               textvariable=start_date)
cal.place(x=400, y=20)
start_date.trace('w', lambda *_, f=f2, d=start_date, i=intervenantVar: get_start_date(*_, frame=f, d=d, p=p,
                                                                                      i=i))

f1.pack(fill='both', side=TOP)

# print("start loading plannings")
load_all_plannings()
# print("plannings loaded")
# print_planning(df_all_plannings)

intervenantEntry.bind("<Return>", lambda e: call_filter_planning(f2, cal, intervenantEntry))

promoLabel = Label(f1, text="Promotions")
promoLabel.place(x=250, y=0)
promoListbox = Listbox(f1, height=len(df_plannings), width=20, selectmode=MULTIPLE)
promoListbox.place(x=250, y=20)

i = 0
for p in df_plannings.keys():
    promoListbox.insert(i + 1, p)
    filter_param['promos'][p] = True
    i = i + 1

f1.pack(fill='both', side=TOP)

c1 = partial(call_filter_planning, f2, cal, intervenantEntry)
filterButton = Button(f1, text="filtrer plannings", activebackground='green', command=c1)
filterButton.place(x=0, y=40)

c2 = partial(prepare_email)
prepareEmailButton = Button(f1, text="preparer email", activebackground='blue', command=c2)
prepareEmailButton.place(x=120, y=40)

c3 = partial(load_all_plannings)
reloadButton = Button(f1, text="recharger plannings", activebackground='blue', command=c3)
reloadButton.place(x=0, y=70)

c4 = partial(find_collision, f2)
findCollisionButton = Button(f1, text="détecter collisions", activebackground='red', command=c4)
findCollisionButton.place(x=120, y=70)

promoListbox.select_set(0, END)
filter_plannings(f2)
statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " lignes selectionnées")
get_intervenants()

top.mainloop()
