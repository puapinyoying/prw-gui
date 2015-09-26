#! usr/bin/env python

import csv # import CSV tools
import sys
import os
import re
import glob
from openpyxl import Workbook # import python Excel parser

NAME_REGEXP = r'(S\d+_(\w+?)_.+)\.csv'

numArgs = len(sys.argv)
if numArgs == 3:
    # create the names for all the files to open
    if os.path.isdir(sys.argv[1]):
        os.chdir(sys.argv[1]) 
        csvFilesInDir = sorted(glob.glob('*.csv'))
        excelFileName = sys.argv[2] + '.xlsx'
        csvNamesNoExt = []
        sheetNames = []
        
        for fullname in csvFilesInDir:
            nameSearchObj = re.search(NAME_REGEXP, fullname)
            csvNamesNoExt.append(nameSearchObj.group(1))
            sheetNames.append(nameSearchObj.group(2))
         
        # Create an Excel workbook to house this stuff
        wb = Workbook()
        
        # Create a new sheets
        for i, name in enumerate(sheetNames):
            wb.create_sheet(index=i, title=name)
    
        # populate sheets
        for i, csvFileName in enumerate(csvFilesInDir):
            with open(csvFileName, 'rb') as csvFile:
                csvFileReader = csv.reader(csvFile, delimiter=',', quotechar='"')
                ws = wb.worksheets[i]
                for row in csvFileReader:
                    ws.append(row)
     
        # Save the excel file
        wb.save(excelFileName)
     
        print "Complete. Data output to: %s" % (excelFileName)
    else:
        print "Check your args"
else:
    print 'USAGE'