## Excel Comparison Tool

####Python tool using OpenPyXL, DiffLib to iteratively compare Excel 2007+ (ie, xlsx) worksheets and write the differences to file.

The program uses *<a href="https://openpyxl.readthedocs.org/en/latest/" target="_blank">OpenPyXL</a>* and *<a href="https://docs.python.org/2/library/difflib.html" target="_blank">DiffLib</a>* to read in an Excel workbook and compare multiple worksheets from within. The program performs the following:
- concatenates all columns from each row in the 'pre' worksheet and performs a lookup on the 'post' worksheet
- any differences are summarised in the console and ouput to file including the two 'next closest matches' 
- the program then iteratively compares ```n-1``` columns to determine where the main differences reside.
  
To use the program make sure your workbook and worksheets satisfy the following checklist:

|Checklist Item										| Workbook or Worksheet| Variable Name |
|:---------------------------------------------------|:-------|:---------------------|
|workbook name is defined as a string in the 'ready' variable  							| Workbook	|	```'ready'```|
|worksheets for comparison are defined in the list variables: 'sheet_list_old', 'sheet_list_new' - you can do multiple, but ensure they are in sequence						| Worksheet	| ```'sheet_list_old', 'sheet_list_new'```  |
|ensure there are the same number of columns for comparison in the 'pre' and 'post' worksheets and in the same order. Columns do not need to be sorted.			| Worksheet	|	na  |


In the 'utilities' folder there are a few older versions including 'pandas' based variation. There are also variations to compare different workbooks, which include a MD5 Checksum. The version currently only support identically sized and sorted worksheets/workbooks.

  
This has been developed/tested on Python 2.7 and Excel 2010 xlsx workbooks.


![OpenPyXL](https://openpyxl.readthedocs.org/en/latest/_static/logo.png)
