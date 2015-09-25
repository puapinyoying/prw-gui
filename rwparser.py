import re # import regular expression library
import sys # import OS/system level tools
import csv # import CSV tools
import os
from PySide import QtCore, QtGui
from mainwindow import Ui_MainWindow
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook # import python Excel parser

# Regular Expressions and static (unchanging) variables 
SAMPLE_NAME_REGEXP = r'([\w \/]+) Turns Data'
FILE_NAME_REGEXP = r'(.+)\..+'

class RwParser:
    def getFileNamesAndDir(self, fileNameTuple):
        """Gets full file path, directory name (full path minus filename), full
        filename and filename without the .asc extension from fileNameTuple 
        QtGui.QFileDialog.getOpenFileName()"""
        
        fullFilePath = fileNameTuple[0]
        fullFileName = os.path.basename(fullFilePath)
        dirName = os.path.dirname(fullFilePath)
        # Use a regular expression to match the filename
        nameSearchObj = re.search(FILE_NAME_REGEXP, fullFileName)       
        # Capture part without extension
        fileNameNoExt = nameSearchObj.group(1)
        return fullFilePath, dirName, fullFileName, fileNameNoExt
    
    def mkParsedFolder(self, dirName, fileNameNoExt):
        """Makes a new folder for parsed data"""
        newFolderName = fileNameNoExt +'_parsed'
        newFolderPath = os.path.join(dirName, newFolderName)
        if not os.path.exists(newFolderPath):
            os.makedirs(newFolderPath)
        return newFolderPath
    
    def checkFileHeader(self, fullFilePath, MAIN_HEADER):
        """Function to check the file's main header. If header right return 
        True, else False."""
        # Try a search and find on the header row to match HEADER_STRING. If return
        # error, print feedback and quit program.
        #Move to the next (first) row of data and assign it to a variable
        with open(fullFilePath, 'rb') as csvFile:
            # Turns the open file into an object we can use to pull data from
            ascFileReader = csv.reader(csvFile, delimiter=",", quotechar='"')
    
            # Check first file header and assign to variable
            try:
                firstHeader = ascFileReader.next()
                searchHeader = re.search(MAIN_HEADER, firstHeader[0])
                searchResult = searchHeader.group(0) # group(0) returns whole string
                print searchResult
                return True
            except (csv.Error, AttributeError): # bring up pop msg if wrong file type
                msgBox = QtGui.QMessageBox()
                QtGui.QPlainTextEdit.LineWrapMode
                msgBox.setText("ERROR: This file is in the wrong format or is the wrong file type.")
                msgBox.exec_()
                return False
          
            
    def checkSampleHeader(self, SAMPLE_NAME_REGEXP, rowOfData):
        """Function to check the second header and formats header for distance csv
        Returns raw header and reformatted header"""
        rowLength = len(rowOfData)
    
        # Use the modulo (%) operator which yeilds remainder after dividing by a
        # number. Sample data should be in triplicate columns. If it's not, quit.
        # Else, reformat the header and return formated one
        if rowLength % 3 != 0:
            print "ERROR: Sample data are not in triplicate columns."
            sys.exit(0) # quit the program
        else:
            # For distance file, reformat header
            distHeaderRow = ["Date", "Time"]
    
            # Loop through the columns of data starting at col 3, until end. Step 3
            # at a time to get col containing Turn Data for each sample only.
            # Remember python lists start at 0, 1, 2, ...
            # I reformat the column order here to remove duplicate time and dates
            # and add a meters per minute one.
            for i in xrange(2, rowLength, 3):
                sampleHeader = rowOfData[i]
    
                # Add the Turns Data sample header as is to the row
                #distHeaderRow.append(sampleHeader)
    
                # Capture the sample name (without 'Turn Data')
                searchObj = re.search(SAMPLE_NAME_REGEXP, sampleHeader)
                searchResult = searchObj.group(1)
    
                # Append meters/min to end of name
                sampleDistHeader = searchResult + ' meters/min'
    
                # Add this new meters/min sample header as a new column
                distHeaderRow.append(sampleDistHeader)
            # Return True for header checked, and the newly created header
            return distHeaderRow
    
    def firstPass(self, fullFilePath, fileNameNoExt, newFolderPath):
        """Runs through file, checks headers.  If correct generates two csvs of
        sample summary and raw sample data. Also returns unique dates, start and
        end times for filtering in later steps and a new header for the distance
        conversion file"""
        
        # Set a couple of variables for filtering time
        uniqueDates = []
        startTime = ''
        endTime = ''
        
        # keep track of sample rows
        firstSampRow = True

        # Grab the entire file name and the name without the (.asc) extension
 
        # Make a few file names
        miceCsvName = os.path.join(newFolderPath, fileNameNoExt + '_mice.csv')
        rawCsvName = os.path.join(newFolderPath, fileNameNoExt + '_rawData.csv')
        distCsvName = os.path.join(newFolderPath, fileNameNoExt + '_distData.csv')
        
    
        # Open a connection to the provided csv file to read from
        with open(fullFilePath, 'rb') as csvFile:
            # Turns the open file into an object we can use to pull data from
            ascFileReader = csv.reader(csvFile, delimiter=",", quotechar='"')
    
            # Create a mice data csv file
            with open(miceCsvName, 'wb') as miceOutFile:
                miceFileWriter = csv.writer(miceOutFile)
    
                # Create a raw data csv file
                with open(rawCsvName, 'wb') as rawOutFile:
                    rawFileWriter = csv.writer(rawOutFile)
    
                    # Create a calculated distance (meters/min) file
                    with open(distCsvName, 'wb') as distOutFile:
                        distFileWriter = csv.writer(distOutFile)
                        
                        # Make sure to check the data header too
                        checkedSampleHeader = False
    
                        # Iterate through every line (row) of the original file
                        for row in ascFileReader:
                            # Put all mouse summary data into the mice sheet, data
                            # should be in less than 3 columns
                            if len(row) < 3:
                                miceFileWriter.writerow(row)
                        
                            # Real data should have at least 3 columns
                            else:
                                # Check second header, is it a multiple of 3?
                                if not checkedSampleHeader:
                                    distHeader = self.checkSampleHeader(
                                        SAMPLE_NAME_REGEXP, row)
    
                                    checkedSampleHeader = True
                                    
                                    # Raw data file takes header row as is
                                    rawFileWriter.writerow(row)
                                           
                                    # Distance data file takes the formatted one
                                    distFileWriter.writerow(distHeader)
    
                                # Once the header is good, parse the data and
                                # calculate the meters/min
                                else:
                                    if firstSampRow:
                                        startTime = row[1]
                                        firstSampRow = False
                                    rawFileWriter.writerow(row)
                                    distanceRow = [row[0], row[1]]
                                    for i in xrange(2, len(row), 3):
                                        sampleData = row[i]
                                        #distanceRow.append(sampleData)
                                        meterData = float(sampleData) * 0.361
                                        distanceRow.append(meterData)
                                    distFileWriter.writerow(distanceRow)
                                    
                                    # will be overwritten until reaches last row
                                    endTime = row[1]
                                    
                                    # grab unique dates
                                    if row[0] not in uniqueDates:
                                        uniqueDates.append(row[0])
        
        return uniqueDates, startTime, endTime, distCsvName
                
if __name__ == '__main__':
        # File variables
        rwparser = RwParser()
        