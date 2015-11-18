#!/usr/bin/env python
# -*- coding: utf-8 -*-

## ADAPATED VERSION FOR ROW by ROW and then CELL by CELL Comparison ##

import openpyxl
from openpyxl.cell import get_column_letter, column_index_from_string
from openpyxl import load_workbook
import hashlib
import csv
import pprint as pp
import itertools
from itertools import cycle
import re
from string import punctuation


ready = 'Helix_Case_PY.xlsx'

sheet_list_old = ['Helix_Case_DCW']
sheet_list_new = ['Helix_Case_AUDIT']


#function to rerun lookup less one attribute
def list_stripper(row, n):
    global row_list_cat
    row_list_cat = []

    [row_list_cat.append([''.join([cell.internal_value for cell in row[0:n]]), cell.coordinate, n])]

    return row_list_cat
        
    
#function to write CSV file
def writer_csv(output_list):
    
    #uses group name from URL to construct output file name
    file_out = "DCW_Compare_{dcw}.csv".format(dcw = ready.rsplit('.',2)[0])
    
    with open(file_out, 'w') as csvfile:
        col_labels = ['Compare_ID', 'Columns_Compared', 'Lookup_String', 'DCW_CellRef', 'Closest_Match_Audit']
        
        writer = csv.writer(csvfile, lineterminator='\n', delimiter=',', quotechar='"')
        newrow = col_labels
        writer.writerow(newrow)
        
        for i in output_list:
            
            newrow = i['compare_id'], i['columns_compared'], i['lookup_value_DCW'], i['cell_ref_dcw'], i['lookup_value_AUDIT']
            writer.writerow(newrow)      

    
#iterate through sheets and identify cells that do not match 
def sheet_checker(ready):    
    
    output_list = []

    #load workbooks for DCW and Audit Report
    wb_all = openpyxl.load_workbook(ready, use_iterators=True, data_only=True)

 
    for i, j in zip(sheet_list_new, sheet_list_old):

        ws_new = wb_all.get_sheet_by_name(i)
        ws_old = wb_all.get_sheet_by_name(j)

        row_new_list = []
        row_old_list = []
        
        compare_new_list = []
        compare_old_list = []
        
        #"map" with paramter 'None' ensures that lists of different length can be handled
        for row_new, row_old in map(None, ws_new.iter_rows(), ws_old.iter_rows()):
            

            #check this: only row_old required i suspect
            if row_new is not None and row_old is not None:

                #this will define how many column stripping cycles to run
                n_new = len(row_new) + 1
                n_old = len(row_old) + 1
                
                #create lists of 'cascading' concatenations
                for n in range(1, n_new):

                    compare_new = list_stripper(row_new, n)
                    compare_new_list.append(compare_new)

                    compare_old = list_stripper(row_old, n)
                    compare_old_list.append(compare_old)

          
        #create list for only concatenated strings from NEW
        list_lookup_new = []
        for p in compare_new_list:
            for q in p:
                list_lookup_new.append(q[0])
        
        #output results for lookups
        x=1
        for e in compare_old_list:
            mismatch_dict = {}
            for f in e:
                if f[0] not in list_lookup_new:
                    #print str(x) + ' - Columns Compared: ' + str(e[0][2]) + ' - Value: ' + str(e[0][0]) + ' - DCW CellRef: ' + str(e[0][1])
                    
                    #regex pattern to find closest match          
                    pattern = re.compile(f[0].format(re.escape(punctuation)), re.IGNORECASE)
                    
                    mismatch_dict = {
                        'compare_id' : str(x),
                        'columns_compared' : str(e[0][2]),
                        'lookup_value_DCW' : str(e[0][0]),
                        'lookup_value_AUDIT' : ', '.join(set(filter(None, [pattern.search(z).group() if pattern.search(z) is not None else "" for z in list_lookup_new]))),
                        'cell_ref_dcw' : str(e[0][1])
                        }
                    
                    output_list.append(mismatch_dict)
                    
                x += 1

        writer_csv(output_list)
        
       
        
        
if __name__ == "__main__":
    sheet_checker(ready)
      
    
