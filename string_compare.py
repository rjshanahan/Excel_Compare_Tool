#compares long text strings - imports them into a list

import csv
import pprint as pp


with open('DCW_Helix_Case.csv', 'rb') as f:
    reader = csv.reader(f)
    DCW_list = list(reader)
    
    

with open('Audit_Helix_Case.csv', 'rb') as f:
    reader = csv.reader(f)
    Audit_list = list(reader)
    
 
DCW_check = []
Audit_check = []

    
if len(DCW_list) != len(Audit_list):
    print 'warning! the lists are of different lengths - this may indicate duplicate'

    
for i in DCW_list:
    if i not in Audit_list:
        DCW_check.append(i)
    
    else:
        pass
    
print 'The following item was not found in the AUDIT: ' 
pp.pprint(DCW_check)


    
print '\n\n\n'

for i in Audit_list:
    if i not in DCW_list:
        Audit_check.append(i)
    
    else:
        pass
    
print 'The following item was not found in the DCW: ' 
pp.pprint(Audit_check)
    
