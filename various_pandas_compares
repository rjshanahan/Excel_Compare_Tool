import pandas as pd
import numpy as np

#workbook to be imported
ready = 'Helix_Case_PY.xlsx'

#import DCW tab
df_DCW = pd.read_excel(ready,
                       'Helix_Case_DCW',
                       skipinitialspace=True)

#import AUDIT tab
df_AUDIT = pd.read_excel(ready,
                         'Helix_Case_AUDIT',
                         skipinitialspace=True)

#functions to clean values on load
def strip(text):
    try:
        return text.strip()
    except AttributeError:
        return text

def make_int(text):
    return int(text.strip('" '))


#inspect dataframes
#df_DCW.head(5)
#df_AUDIT.head(5)

#create lists of column headers to concatenate
df_DCW_cols = list(df_DCW.columns.values[0:2])
df_AUDIT_cols = list(df_AUDIT.columns.values[0:2])

#add 'concatenate' attruibute
df_DCW['concatenate'] = df_DCW[df_DCW_cols].apply(lambda x: ''.join(x), axis=1)
df_AUDIT['concatenate'] = df_AUDIT[df_AUDIT_cols].apply(lambda x: ''.join(x), axis=1)


#add 'source' attribute
df_DCW['source'] = 'DCW'
df_AUDIT['source'] = 'AUDIT'


#'merge' dataframes on the 'concatenate' column
df_MERGE = pd.merge(df_DCW, 
                    df_AUDIT, 
                    on=['concatenate'], 
                    #on=df_DCW_cols, 
                    how='outer')


.Alpha Mannosidase activity PlasmaPrimary.Alpha Mannosidase activity PlasmaA MAN PLASMAGeneral ChemistryLab - Gen LabX
.Alpha Mannosidase activity PlasmaPrimary.Alpha Mannosidase activity PlasmaA-MAN PLASMAGeneral ChemistryLab - Gen LabX



#df_MERGE.head(100)

#df_MERGE = df_MERGE.dropna()

#df_MERGE = df_MERGE[df_MERGE.isnull()]

df_MERGE = df_MERGE.fillna('MISSING')


#df_MERGE = df_MERGE[df_MERGE.isnull().any(axis=1)]

df_MERGE



with pandas.io.excel.ExcelWriter(path=Path, engine="xlsxwriter") as writer:
   sheet = writer.book.worksheets()[0]
   sheet.write(x, y, value, format) #format is what determines the color etc.
