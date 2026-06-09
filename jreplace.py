#! /usr/bin/python
'''
Replace the keywords in a word .docx template by the content
of the columns with the same names in a specific row(s) of an excel sheet.

syntax:  jreplace  xlsx  docx  rownumber1 rownumber2 ...

The keywords used in the word template must correspond exactly to columns
in the first row of the excel sheet. 

The first sheet in the excel document will be used.

The output file(s) will have the same name as the template, but with a
tag indicating the associated row number.

The docx template can be replaced by a directory tree with multiple docx
documents, in which case the search/replace will apply to all the documents, 
and a new directory tree will be created.

Carlos Allende Prieto, October 2016
Carlos, December 2016, fixed it to work with keywords within tables and 
                       handle page breaks
Carlos, December 2016, upgraded to work with dir trees containing templates
'''

import sys
import os
import openpyxl
from jreplace_core import jreplace_file, jreplace_dir


def info():
    """
    inform users on syntax
    """
    print('syntax:  jreplace  xlsx  docx/dirtree  rownumber1 rownumber2 ...')
    return


def prechecks():
    """
    carry out basic checks on the command-line arguments
    """
    if len(sys.argv) < 4:
        print('ERROR: at least 3 arguments are required')
        info()
        exit()
    if int(sys.argv[3]) < 2:
        print('ERROR: the 3rd argument (rownum) must be >=2')
        info()
        exit()
    if sys.argv[1][-4:] != 'xlsx':
        print(sys.argv[1][-4:])
        print('ERROR: the 1st argument must be the name of the xlxs file')
        info()
        exit()

    if os.path.isfile(sys.argv[2]):
        if sys.argv[2][-4:] != 'docx':
            print('ERROR: the 2nd argument must be the name of the docx file')
            info()
            exit()
    elif os.path.isdir(sys.argv[2]):
        pass
    else: 
        print("ERROR: template " + sys.argv[2] + " is neither a file nor a directory")
        info()
        exit()


if __name__ == "__main__":

    # carry out basic checks
    prechecks()

    # assign parameters to variables
    workbook = sys.argv[1]
    template = sys.argv[2]
    rownumbers = sys.argv[3:]

    # open excel sheet
    wb = openpyxl.load_workbook(workbook)
    sheetnames = wb.get_sheet_names()
    sheet = wb.get_sheet_by_name(sheetnames[0])

    # loop over rownumbers
    for row in rownumbers:
        rownum = int(row)
  
        if rownum > sheet.max_row:
            print('ERROR: requested rownumber=' + row + ' but the excel sheet has only ' + \
                   str(sheet.max_row) + ' rows') 
            continue

        tag = ("_%04d" % rownum)

        if os.path.isfile(template):
            jreplace_file(sheet, template, rownum, tag)
            print('\n-- Writing output to ' + os.path.basename(template)[0:-5] + tag + '.docx')
        elif os.path.isdir(template):
            jreplace_dir(sheet, template, rownum, tag)
            print('\n-- Created output directory: ' + os.path.basename(template) + tag)
        else: 
            print("ERROR: template " + template + " is neither a file nor a directory")
            break
