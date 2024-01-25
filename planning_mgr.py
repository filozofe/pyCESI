# doc: https://pythonhosted.org/xlrd3/
# pip install babel pandas tkinter pywin32 pandastable pretty_html_table tkcalendar
# installer Python pour Windows 64bits: https://www.python.org/downloads/
# python location: $env:USERPROFILE\AppData\Local\Programs\python\python.exe
# C:\Users\phofmann\AppData\Local\Programs\python\Python311\python.exe -m venv venv
# .\venv\Scripts\Activate.ps1
# python.exe -m pip install --upgrade pip
# pip install babel pandas pywin32 pandastable pretty_html_table tkcalendar openpyxl numpy
 

titre = "gestion plannings v1.1 25/1/2024"

import codecs
import babel
import locale
import os
from functools import partial
from tkinter import *


import pandas as pd
import numpy as np
import win32com.client
from pandastable import Table, config 
from pretty_html_table import build_table
from tkcalendar import Calendar

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
df = '%a %d/%m/%Y'
# locale.setlocale(locale.LC_TIME, 'fr_FR')
# locale.setlocale(locale.LC_ALL, 'fr')




recipient = ''
signature_name = 'new'

# Permanently changes the pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# promos = ('ASR1 23-24', 'ASR2-FFASR 23-24', 'MAALSI 23-25', 'ASR2-FFASR 24-25')

planning_file = r"\\tldata\utilisateurs\Activité Alternance\Activité informatique\PROMOS\plannings_informatique.xlsx"

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
        #print("loading " + plannings.iloc[i]['promo'])
        statusLabel.config(text="chargement planning " + str(plannings.iloc[i]['promo']) + ": " + str(
            plannings.iloc[i]['fichier']) + "                                    ")
        top.update_idletasks()
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
            df2['Pilote'] = plannings.iloc[i]['pilote']
            df2.replace(np.nan, '', inplace=True)
            df2 = df2[df2['Réf.'] != '']
            #    print(df_all_plannings[df_all_plannings['Intervenants'] == 'FOLIO Brice'])
            df2['Jour'] = df2['Date'].dt.strftime('%a %d/%m/%Y')
            df2.rename({'AM-PM':'AM_PM'},inplace=True,axis='columns')         #remplace AM-PM par AM_PM
            #normalisation des demies journées
            df2.replace(['AM','am','matin','Matin'],'matin',inplace=True)
            df2.replace(['pm','PM','aprem','Aprem'],'aprem',inplace=True)
            #cast to correct types
            #df2 = df2.astype({'Charge': 'float'}).dtypes
            #remove leading and trailing spaces
            df2['Intervenants']=df2['Intervenants'].apply(lambda x: x.strip())
            df2.sort_values('Date')
            df2.reset_index(inplace=True, drop=True)
            df_plannings[plannings.iloc[i]['promo']] = df2
    top.update_idletasks()
    statusLabel.config(text=str(len(plannings)) + " plannings chargés ")
    top.update_idletasks()
    return



#pivot table avec les intervenants et les noms de promo
def get_intervenants2(frame):
    result_df = pd.DataFrame(columns=['Intervenants', 'Pilote'])
    
    for df in df_plannings.values():
        result_df = result_df._append(df[['Intervenants','Pilote']])
    result_df.sort_values('Intervenants')

    pivot_df=pd.pivot_table(result_df,values='Pilote',index=['Intervenants'],columns=['Pilote'],aggfunc=len,fill_value='')
    pivot_df['Intervenant']=pivot_df.index
    cols = pivot_df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    pivot_df=pivot_df[cols]
    pt = Table(frame,
               dataframe=pivot_df,
               showtoolbar=True, showstatusbar=True)
    # set some options
    options = {'colheadercolor': 'blue', 'floatprecision': 1, 'fontsize': 8, 'cellwidth': 60}
    config.apply_options(options, pt)
    top.update_idletasks()
    pt.show()
    top.update()
    top.update_idletasks()
    pt.redraw()
    return



def get_pilotes():
    pilotes = set()
    for df in df_plannings.values():
        pilotes = pilotes.union(set(df['Pilote'].tolist()))
    #print(pilotes)
    return pilotes


def filter_plannings(frame):
    global df_filtered_plannings
    result_df = pd.DataFrame()

    # delete all rows
    df_filtered_plannings = df_filtered_plannings.head(0)

    # print("date ", filter_param['date'])
    # print("promos", filter_param['promos'])
    # print('intervenant', filter_param['intervenant'])
    #construire la liste de filtrage des pilotes
    pilotes = []
    for key, values in filter_param['pilote'].items():
        if values:
            pilotes.append(key)
    for p, df in df_plannings.items():
        if filter_param['promos'][p] == True:
            df_filtered_plannings = df_filtered_plannings._append(df.loc[
                                                                      (df['Intervenants'].str.contains(
                                                                          filter_param['intervenant'])) &
                                                                      (df['Pilote'].isin(pilotes)) &
                                                                      (df['Date'] >= pd.to_datetime(filter_param['date'],
                                                                                                   dayfirst=True))
                                                                      ])
    if filter_param['confirmé'] == validationFilterOptions[1]:
        #que les confimé
        #df_filtered_plannings=df_filtered_plannings.drop[df_filtered_plannings[df_filtered_plannings['Confirmation']!=''].index]
        df_filtered_plannings=df_filtered_plannings[df_filtered_plannings.Confirmation != '']
    elif filter_param['confirmé'] == validationFilterOptions[2]:
        #que les non confirmé
        #df_filtered_plannings=df_filtered_plannings.drop[df_filtered_plannings[df_filtered_plannings['Confirmation']==''].index]
        df_filtered_plannings=df_filtered_plannings[df_filtered_plannings.Confirmation == '']
    else:
        pass
    df_filtered_plannings.sort_values('Date', inplace=True)
    df_filtered_plannings.reset_index(inplace=True, drop=True)
    pt = Table(frame,
               dataframe=df_filtered_plannings[
                   ['Pilote', 'Jour', 'AM_PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module', 'Confirmation']],
               showtoolbar=False, showstatusbar=True)
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

    filter_plannings(frame)

    df_filtered_plannings['collision'] = df_filtered_plannings.duplicated(subset=['Jour', 'Intervenants', 'AM_PM'],keep=False)
    #df_filtered_plannings['collision'] = df_filtered_plannings.duplicated(subset=['Jour', 'Intervenants'],keep=False)
                                                                          
    df_filtered_plannings = df_filtered_plannings[df_filtered_plannings.collision == True]
    # remove from duplicate the following 'Intervenants'
    exclude = ['', '?', 'Pilote', 'pilote', 'Autonomie']
    df_filtered_plannings = df_filtered_plannings[~df_filtered_plannings.Intervenants.isin(exclude)]
    df_filtered_plannings.sort_values('Date', inplace=True)
    df_filtered_plannings.reset_index(inplace=True, drop=True)

    pt = Table(frame,
               dataframe=df_filtered_plannings[
                   ['Pilote','Jour', 'AM_PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module', 'Confirmation']],
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
            p[['Jour', 'AM_PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module', 'Confirmation']].to_string(
                index=False,
                justify='left'))


def display_plannings():
    ### https://pypi.org/project/pretty-html-table/
    html_beautiful_table = build_table(
        df_filtered_plannings[['Jour', 'AM_PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
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
        df_filtered_plannings[['Jour', 'AM_PM', 'Charge', 'Intervenants', 'Promo', 'Réf.', 'Module']],
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
def call_filter_planning(frame):
    date = date = cal.get_date()
    if date == '':
        date = today
    filter_param['date'] = date
    filter_param['confirmé'] = selectedOption.get()
    filter_param['intervenant'] = intervenantVar.get()
    # set all promo to false
    for p in filter_param['promos'].keys():
        filter_param['promos'][p] = False
    for i in promoListbox.curselection():
        # print(promoListbox.get(i))
        filter_param['promos'][promoListbox.get(i)] = True

    for p in filter_param['pilote'].keys():
        filter_param['pilote'][p] = False
    for i in piloteListbox.curselection():
        # print(promoListbox.get(i))
        filter_param['pilote'][piloteListbox.get(i)] = True

    displayFilterParam()
    filter_plannings(frame)
    print( df_filtered_plannings.info())
    statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " lignes selectionnées")
    # display_plannings()
    return

validationFilterOptions = ["confirmé et non confirmé","confirmé","non confirmé"]

################# main
filter_param = {
    'date': Calendar.date.today().strftime("%d/%m/%y"),
    'promos': {},
    'intervenant': '',
    'pilote': {},
    'confirmé': "tous"         #si False on ne filtre pas le entrees qui on un valeur dans la colonne validation
}





# display les paramettres de filtre
def displayFilterParam():
    string = str()
    string = "Date début: " + filter_param['date'] + '\n'
    string = string + "confirmé: " + filter_param['confirmé'] + '\n'
    string = string + "Intervenant: " + filter_param['intervenant'] + '\n'
    string = string + "Pilote: "
    for key, value in filter_param['pilote'].items():
        if value:
            string = string + key + ", "
    string = string + "\nPromos: "
    for key, value in filter_param['promos'].items():
        if value:
            string = string + key + " "
    filter_param_text.set(string)


top = Tk()
top.title(titre)
top.geometry("1200x800")
# f1 en haut : les commandes
# f2 bas droite: la table
# f3: a gauche: la liste des promos
f1 = Frame(top, height=220)  # for the buttons and the entries
f1.pack(fill='both', side=TOP)
f4 = Frame(top, height=20)  # pour les status
f4.pack(fill='x', side=BOTTOM)
f3 = Frame(top, width=110)  # pour la liste des promos
f3.pack(fill='y', side=LEFT)
f2 = Frame(top, width=1090, height=560)  # for the tables
f2.pack(fill='both', expand=True)

intervenantLabel = Label(f1, text="Intervenant")
intervenantLabel.place(x=1, y=10)
intervenantVar = StringVar()
intervenantEntry = Entry(f1, textvariable=intervenantVar)
intervenantEntry.place(x=80, y=10)

statusLabel = Label(f4, text="starting")
statusLabel.place(x=1, y=0)

start_date = StringVar(f1, Calendar.date.today().strftime("%d/%m/%y"))

Label(f1, text="Date début").place(x=400, y=0)
cal = Calendar(f1,
               selectmode='day',
               date_pattern='dd/mm/yyyy',
               year=today.year, month=today.month, day=today.day,
               locale='fr',
               # start_date=Calendar.date.today().strftime("%d/%m/%y"),
               textvariable=start_date)
cal.place(x=400, y=20)
#start_date.trace('w', lambda *_, f=f2, d=start_date, i=intervenantVar: get_start_date(*_, frame=f, d=d, p=p, i=i))

start_date.trace('w', lambda *_, f=f2 : call_filter_planning(f))

Label(f1, text="filtre").place(x=670, y=0)
filter_param_text = StringVar()
filter_param_text.set("Filtre")
filterParamLabel = Label(f1, textvariable=filter_param_text, anchor=W, justify=LEFT, wraplength=500)
filterParamLabel.place(x=670, y=20)
intervenantEntry.bind("<Return>", lambda e: call_filter_planning(f2))

promoLabel = Label(f3, text="Promotions")
promoLabel.place(x=10, y=0)
promoListbox = Listbox(f3, height=len(df_plannings), width=15, selectmode=MULTIPLE, exportselection=False)
promoListbox.place(x=0, y=20)
promoListbox.bind('<<ListboxSelect>>', lambda e: call_filter_planning(f2))

Label(f1, text="Pilotes").place(x=250, y=0)
piloteListbox = Listbox(f1, width=15, selectmode=MULTIPLE, exportselection=False)
piloteListbox.place(x=250, y=20)
piloteListbox.bind('<<ListboxSelect>>', lambda e: call_filter_planning(f2))

c1 = partial(call_filter_planning, f2)
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

c5 = partial(promoListbox.select_set, 0, END)
selectAllPromoButton = Button(f1, text="select. tte promos", command=c5)
selectAllPromoButton.place(x=0, y=100)

c6 = partial(promoListbox.select_clear, 0, END)
deSelectAllPromoButton = Button(f1, text="deselect. tte promos", command=c6)
deSelectAllPromoButton.place(x=120, y=100)

selectedOption=StringVar()
r1=Radiobutton(f1,text=validationFilterOptions[0],value=validationFilterOptions[0], variable=selectedOption,command=c1)
r2=Radiobutton(f1,text=validationFilterOptions[1],value=validationFilterOptions[1], variable=selectedOption,command=c1)
r3=Radiobutton(f1,text=validationFilterOptions[2],value=validationFilterOptions[2], variable=selectedOption,command=c1)
r1.place(x=0, y=160)
r2.place(x=0, y=180)
r3.place(x=0, y=200)
selectedOption.set(validationFilterOptions[0])

c7 = partial(get_intervenants2, f2)
showIntervenantButton=Button(f1, text="lister Intervenants", command=c7)
showIntervenantButton.place(x=0,y=130)



def load():
    load_all_plannings()

    # load promo list box
    i = 0
    for p in df_plannings.keys():
        promoListbox.insert(i + 1, p)
        filter_param['promos'][p] = True
        i = i + 1
    promoListbox.select_set(0, END)

    # load pilote listbox
    i = 0
    pilotes = get_pilotes()
    for p in pilotes:
        piloteListbox.insert(i + 1, p)
        filter_param['pilote'][p] = True
        i = i + 1
    piloteListbox.select_set(0, END)

    filter_plannings(f2)
    statusLabel.config(text=str(df_filtered_plannings.shape[0]) + " lignes selectionnées")
    displayFilterParam()


top.after(1000, load)
top.mainloop()
