#doc: https://pythonhosted.org/xlrd3/



import pandas as pd

# Open the Workbook
df = pd.read_excel(r"C:\Users\phofmann\OneDrive - Cesi\Bureau\ASR1 suivi.xlsx",
                   sheet_name='ASR1')

df.info(verbose=True)
print(df[['nom','prenom','PortablePersonnel','mail cesi']])