import textwrap
import re  # import regular expression library
import sys  # import OS/system level tools
from os import path
from PySide import QtCore, QtGui
from mainwindow import Ui_MainWindow
from rwparser import RwParser
from datetime import datetime, timedelta




TUTORIAL = """Tutorial

This program parses mouse running wheel data that has been output into an ASCII (.asc) file directly obtained from the Vitalview Software.    The input file    should have two headers. The first will be the the over all file header that says 'Experiment Logfile:' followed by the date. Second, the data header below in comma delimited format (csv).

The program will output a csv file for each of the following:

        1) Mice data - unmanipulated summary statistics from ASCII file

        2) Raw data - unmanipulated data from ASCII file

        3) Distance calculated - converted turns data into meters

        4) Time Filtered - option to set start and end date. If unchanged, all data will be used. This sheet also includes a few summary statistics calculated at the bottom rows.

        5) Sum Hourly - minute data is summed into hours per row

        6) Cumulative - data from each hour is compounded onto the previous

        7) Running Streaks - how many consecutive minutes each animal runs per session. This includes all running sessions for that particular animal.

Instructions: 

        1) To start click "Open File" and select a VitalView (.asc) data file.

        2) Once loaded, if you would like to filter the data by start and end date/time use the drop-down boxes and time edit boxes (time format in hh:mm:ss)

        3) The program will automatically output csv files of all the calculations listed above. If you would like to export all the data into a single Excel file check the box (checked by default).

        4) Click "Parse Data" to start."""

# Constant Variables
MAIN_HEADER = 'Experiment Logfile:'

class ParseWindow(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        QtGui.QMainWindow.__init__(self, parent)
        self.setupUi(self)
        
        # File name and path variables
        self.fileNameTuple = ''
        self.fullFilePath = ''
        self.dirName = ''
        self.fullFileName = ''
        self.fileNameNoExt = ''
        self.newFolderPath = ''
        self.distCsvName = ''
        
        # Default dates and times
        self.uniqueDates = []
        self.startTime = ''
        self.endTime = ''
 
        
        # Tool/menu bar clicks
        self.actionQuit.triggered.connect(self.actionQuit_triggered)
        self.actionAbout.triggered.connect(self.actionAbout_triggered)
        self.actionTutorial.triggered.connect(self.actionTutorial_triggered)
        
        # Push buttons
        self.openButton.clicked.connect(self.openButton_clicked)
        self.parseButton.clicked.connect(self.parseButton_clicked)
        
        # Combo boxes
        self.startDateCombo.currentIndexChanged.connect(self.startDateCombo_changed)
        self.endDateCombo.currentIndexChanged.connect(self.endDateCombo_changed)
        
        # Time edit boxes
        self.startTimeEdit.editingFinished.connect(self.startTimeEdit_edited)
        self.startTimeEdit.editingFinished.connect(self.endTimeEdit_edited)
        
        # Check box
        self.excelCheckBox.clicked.connect(self.excelCheckBox_clicked)

    def startDateCombo_changed(self):
        print "what common"
        
    def startTimeEdit_edited(self):
        print "time to time"
            
    def endDateCombo_changed(self):
        print "why man"
        
    def endTimeEdit_edited(self):
        print "end the time"

    def getFileNameTuple(self):
        return self.fileNameTuple
            
    def parseButton_clicked(self):
        t = self.getFileNameTuple()
        msgBox = QtGui.QMessageBox()
        msgBox.setText(t[0])
        msgBox.exec_()
        
    def excelCheckBox_clicked(self):
        print "that box is hot"
        
    def actionAbout_triggered(self):        
        msgBox = QtGui.QMessageBox()
        msgBox.setText(
        """Parse Running Wheel - Gui\n\nVersion 0.1\n\nBy Prech Uapinyoying""")
        msgBox.exec_()
        
    def actionTutorial_triggered(self):        
        msgBox = QtGui.QMessageBox()
        QtGui.QPlainTextEdit.LineWrapMode
        msgBox.setText(TUTORIAL)
        msgBox.exec_()
    
    def actionQuit_triggered(self):
        self.close()
        
    def openButton_clicked(self):
        self.fileNameTuple = QtGui.QFileDialog.getOpenFileName(self,
            "Open VitalView data file (.asc)", "", "ASCII Files (*.asc);;All Files (*.*)")
        
        if self.fileNameTuple[0] != '':
            rwparser = RwParser() # create rwparser class object
            
            # Get filename and Path information
            self.fullFilePath, self.dirName, self.fullFileName, \
                self.fileNameNoExt = rwparser.getFileNamesAndDir(self.fileNameTuple)
            
            # Check the header, if its correct...
            isCorrectHeader = rwparser.checkFileHeader(self.fullFilePath, MAIN_HEADER)
            if isCorrectHeader:
                # Create a new folder to house the parsed data and get path
                self.newFolderPath = rwparser.mkParsedFolder(self.dirName, self.fileNameNoExt)
                
                # Begin first pass - gets all time variables, parses first 3 files
                self.uniqueDates, self.startTime, self.endTime, self.distCsvName = \
                    rwparser.firstPass(self.fullFilePath, self.fileNameNoExt, self.newFolderPath)
                
                # Fill out the combo boxes
                self.startDateCombo.addItems(self.uniqueDates)
                self.endDateCombo.addItems(self.uniqueDates)
                
                # set the endDate combo to the last unique date
                self.endDateCombo.setCurrentIndex(len(self.uniqueDates)- 1)
                
                # Turn time strings into QTime obj and fill out time edits
                qEndTime = QtCore.QTime.fromString(self.endTime, 'hh:mm:ss')
                qStartTime = QtCore.QTime.fromString(self.startTime, 'hh:mm:ss')
                self.startTimeEdit.setTime(qStartTime)
                self.endTimeEdit.setTime(qEndTime)                
                    
                # Allow user to access filters and buttons
                self.filterLabel.setEnabled(1)
                self.filterLine.setEnabled(1)
                self.startDateLabel.setEnabled(1)
                self.startDateCombo.setEnabled(1)
                self.startTimeLabel.setEnabled(1)
                self.startTimeEdit.setEnabled(1)
                self.endDateLabel.setEnabled(1)
                self.endDateCombo.setEnabled(1)
                self.endTimeEdit.setEnabled(1)
                self.excelCheckBox.setEnabled(1)
                self.parseButton.setEnabled(1)
                
                # Change label to give user feedback
                self.changableLabel.setText(self.fullFileName + ' successfully loaded.')
            
            else:
                self.changableLabel.setText(self.fullFileName + ' is not a VitalView file')

                
if __name__ == '__main__':
        # File variables
        fullFilePath = ''
        app = QtGui.QApplication(sys.argv)
        MainApp = ParseWindow()
        MainApp.show()
        sys.exit(app.exec_())
        print fullFilePath
