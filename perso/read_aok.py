# read "C:\Users\phofmann\Dropbox\private\finances\aok2_2022-11-27.ods"

import pandas as pd
import datetime
from pandas_ods_reader import read_ods
import re
from numpy import *
import operator

#odsfile = r"C:\Users\phofmann\Dropbox\private\finances\ok2_2024-01-01.ods"
odsfile = r"D:\data\Dropbox\Dropbox\private\finances\aok2_2024-01-01.ods"

# Permanently changes the pandas settings
""" pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.date_dayfirst', True)
pd.set_option('display.max_colwidth', 80) """

def getData():
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
    #libelle = df["description"].tolist()
    #select only lines with categorisation is not empty
    libelle = df.loc[df["Quoi"] != '',"description"].tolist()
    a_categoriser = df.loc[df["Quoi"] == '',"description"].tolist()
    #print (libelle)             #la liste des libelles
    #categories = df["Quoi"].tolist()
    categories = df.loc[df["Quoi"] != '',"Quoi"].tolist()
    #print (libelle)
    #print(categories)          # la liste des categories
    return libelle,categories,a_categoriser

# cree la liste du vocabulaire dans les libellés, 'bag of words'
def createVocabList(libelle):
    vocabulaire = set([])
    for item in libelle:
        tokens = re.split(r'\W+', item)
        #print (tokens)
        tokens2 = [tok.lower() for tok in tokens if len(tok) > 1]
        #print (tokens2)
        vocabulaire = vocabulaire | set(tokens2)
    #print (vocabulaire)
    return list(vocabulaire)

def parseData(libelle):
    libelle_adj = []
    for item in libelle:
        tokens = re.split(r'\W+', item)
        #print(tokens)
        tokens2 = [tok.lower() for tok in tokens if len(tok) > 1]
        libelle_adj.append(tokens2)
        #print(tokens2)
    return libelle_adj

def bagOfWords2VecMN(vocabulaire,inputSet):
    outputSet = [0] * len(vocabulaire)      #initilise une liste de meme longueur contenant des zero
    for w in inputSet:
            if w in vocabulaire:
                 outputSet[vocabulaire.index(w)] += 1
    return outputSet

def trainNB0(trainData,trainCategorie): 
    categorieListDict = {}
    for i in trainCategorie:
        if i not in categorieListDict.keys():
            categorieListDict[i] = 1
        else:
            categorieListDict[i] += 1
    denom=0
    for key,value in categorieListDict.items():
        denom += categorieListDict[key]
    for key,value in categorieListDict.items():
        categorieListDict[key]=categorieListDict[key]/denom
    
    probDict={}
    numDocs=len(trainData)
    numWords=len(trainData[0])

    for key,value in categorieListDict.items():
        numer = ones(numWords)      #tableau rempli de 1
        numer = numer /10
        denom = 2
        for i in range(numDocs):
            if trainCategorie[i] == key:
                numer += trainData[i]
                denom += sum(trainData[i])
        probDict[key] = log (numer/denom)
    return categorieListDict,probDict

def catgoriseNB(newLibelle,categorieListDict,probDict,vocabulaire):
    tokens = re.split(r'\W+', newLibelle)
    tokens2 = [tok.lower() for tok in tokens if len(tok) > 1]
    vecData = bagOfWords2VecMN(vocabulaire,tokens2)
    outputDict={}
    for key,value in probDict.items():
        outputDict[key]=sum(array(probDict[key])*vecData)+log(categorieListDict[key])
    outputDictSorted=sorted(outputDict.items(),key=operator.itemgetter(1), reverse=True)
    #for i in range(0,4):
        #print (outputDictSorted[i][0],' ',"%.1f" % outputDictSorted[i][1], end='|')
    return outputDictSorted[0][0]



print ("preparing model")
libelle,categorie,a_categoriser=getData()
vocabulaire=createVocabList(libelle)
trainMat=[]
adjLibelle=parseData(libelle)
for i in adjLibelle:
    trainMat.append(bagOfWords2VecMN(vocabulaire,i))
categorieListDict, probDict= trainNB0(trainMat,categorie)
print ("done")


""" for i in a_categoriser:
    print (f"{i:<50}",end=':\t')
    category=catgoriseNB(i,categorieListDict,probDict,vocabulaire)
    print() """

for i in a_categoriser:
    category=catgoriseNB(i,categorieListDict,probDict,vocabulaire)
    print(category)
