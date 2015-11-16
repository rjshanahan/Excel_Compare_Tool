#!/usr/bin/env python
# -*- coding: utf-8 -*-

## ADAPATED VERSION FOR ROW by ROW and then CELL by CELL Comparison ##

import openpyxl
from openpyxl.cell import get_column_letter, column_index_from_string
from openpyxl import load_workbook
import itertools
import hashlib
import csv
import pprint as pp
from itertools import cycle


ready = 'Helix_Case_PY.xlsx'

sheet_list_old = ['Helix_Case_DCW']
sheet_list_new = ['Helix_Case_AUDIT']


#function to generate MD5 checksum for file - NOT USED
def md5(fname):
    hash = hashlib.md5()
    with open(fname, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash.update(chunk)
    return hash.hexdigest()


    
#formatted message re: mismatched cells
def mismatch_printer(sheet, row_new_list, row_old_list):
           
    mismatches = [('Audit Report:', cell_new, 'DCW:',cell_old) for cell_new, cell_old in zip(row_new_list, row_old_list) if cell_new != cell_old]
    
    #mismatch_list = []
    
    #for cell_new, cell_old in zip(row_new_list, row_old_list):
    #    if cell_new != cell_old:
    #        m = ('Audit Report:', cell_new, 'DCW:',cell_old)
    #        mismatches = mismatch_list.append(m)
    
    if len(mismatches) != 0:
        txt_writer(mismatches, sheet)

    else:
        pass
    
    
#function to write mismatches to txt file
def txt_writer(mismatches, sheet):
    
    file_out = "UNIT_TESTING_{DCW}_Errors.txt".format(DCW = sheet)
        
    row1 = 'SHEET: "' + sheet + '" has matching errors - refer to problem cell(s) below' + '\n\n'
    row2 = pp.pformat(mismatches)
    row3 = '\n\n'
               
    text_file = open(file_out, "w")
    text_file.write(str(row1))
    text_file.write(str(row2))
    text_file.write(row3)
    text_file.close()
    
    
    
#function to rerun lookup less one attribute
def list_stripper(row, n):
    
    row_list_cat = []

    [row_list_cat.append(''.join([cell.internal_value for cell in row[0:n - 1]]))]
    
    #print row_list_cat[0:20]
    
    return row_list_cat
        
    
#iterate through sheets and identify cells that do not match 
def sheet_checker(ready):    
    
    global row_new_list
    global row_old_list


    #load workbooks for DCW and Audit Report
    wb_all = openpyxl.load_workbook(ready, use_iterators=True, data_only=True)

 
    for i, j in zip(sheet_list_new, sheet_list_old):

        ws_new = wb_all.get_sheet_by_name(i)
        ws_old = wb_all.get_sheet_by_name(j)

        row_new_list = []
        row_old_list = []
        
        #"map" with paramter 'None' ensures that lists of different length can be handled
        for row_new, row_old in map(None, ws_new.iter_rows(), ws_old.iter_rows()):
            
            if row_new is not None:
                [row_new_list.append([cell_new.coordinate, cell_new.internal_value]) for cell_new in row_new]
        
        
            if row_old is not None:
                [row_old_list.append([cell_old.coordinate, cell_old.internal_value]) for cell_old in row_old]
        
            
            if row_new_list != row_old_list:
                
                if row_new is not None and row_old is not None:
                    
                    n_new = len(row_new)
                    n_old = len(row_old)  
                    
                    for n in range(n_new):
                        if n_new > 2 and n_old > 2:
                            compare_new = list_stripper(row_new, n_new)
                            compare_old = list_stripper(row_old, n_old)
                            
                            
                    x = 1
                    for e in compare_old:
                        if e not in compare_new:
                            print str(x) + '_Columns Compared: ' + str(n_new) + '_' + str(e)
                            x += 1
                           
                        n_new = n_new - 1
                        n_old = n_old - 1
        
        
            #list with cell coordinates and contents

            #mismatch_printer(j, row_new_list, row_old_list)
            #print 'mismatch results have been written to file - please review'
        #else:
            #print 'no issues found'

                
if __name__ == "__main__":
    sheet_checker(ready)
