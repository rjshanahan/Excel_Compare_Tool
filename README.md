## Excel_Compare_Tool

####Python tool using OpenPyXL to compare Excel 2010+ (ie, xlsx) workbooks, worksheets and write the differences to file.

The program uses *<a href="https://openpyxl.readthedocs.org/en/latest/" target="_blank">OpenPyXL</a>* to compare Excel workbooks and worksheets and output any differences to file. The program is summarised as follows:
- take two files - 'old' and 'new'
- perform and compare MD5 Checksum
- if checksums do not match the files are read into memory and a list of each worksheet is created
- for worksheets in the 'old' and 'new', each populated row is read, added to a list and compared
- any differences are written to text file with the cell coordinate and cell internal_value.

A variation is also included that allows users to compare worksheets in the same workbook. In this instance a list of worksheets must firstly be defined by the user.
  
This has been developed/tested on Python 2.7 and Excel 2010 xlsx workbooks.


![OpenPyXL](https://openpyxl.readthedocs.org/en/latest/_static/logo.png)
