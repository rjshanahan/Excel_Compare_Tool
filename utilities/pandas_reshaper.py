#!/usr/bin/env python
# -*- coding: utf-8 -*-

#RESHAPING FUNCTION USING PANDAS

import pandas as pd

#original 'dirty' csv file - import as dataframe
filename = 'Helix_CaseBuild2.csv'
dirty = pd.read_csv(filename)



#list of id variables - copy column names
id_vars = ['Primary Case Name for Ordering','Synonyms']

#list of id variables - copy column names
value_vars = ['Synonym Type 1', 'Synonym Type 2']



#function to reshape
def reshaper(df, id_vars, value_vars):
    
    global clean
    
    #melt dataframe
    clean = pd.melt(df,
                    id_vars,
                    value_vars)
    
    #sort by id variables
    clean =  clean.sort(id_vars)
    
    #write to csv
    file_out = "UNIT_TESTING_{DCW}_Reshaped.csv".format(DCW = filename.replace('.csv',''))
    clean.to_csv(file_out, sep=',', index=False)
    

                
if __name__ == "__main__":
    reshaper(dirty, id_vars, value_vars)
    
  

