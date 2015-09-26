import subprocess
import sys  # import OS/system level tools
import os
from PySide import QtCore, QtGui
from prwlib.mainwindow import Ui_MainWindow
from prwlib.rwparser import RwParser

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

        3) The program will automatically output csv files of all the calculations listed above. Option to export the csv files into a single Excel workbook.

        4) Click "Parse Data" to start."""

# Constant Variables
MAIN_HEADER = 'Experiment Logfile:'

class ParseWindow(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        QtGui.QMainWindow.__init__(self, parent)
        self.setupUi(self)
        
        # Path to csv2excelwb script or executable. Should be in ./tools/
        self.csv2excelwbPath = ''
        # Excel checkbox option
        self.excelCheckState = QtCore.Qt.CheckState.Checked
        
        # Some time constants
        self.QT_MIN = QtCore.QTime.fromString('00:00:00', "hh:mm:ss")
        self.QT_MAX = QtCore.QTime.fromString('23:59:59', "hh:mm:ss")
        
        # File name and path variables
        self.fileNameTuple = ''
        self.fullFilePath = ''
        self.dirName = ''
        self.fullFileName = ''
        self.fileNameNoExt = ''
        self.newFolderPath = ''
        
        # Important when using parse button for 2nd round filter
        self.distCsvName = ''
        
        # Raw dates and times in string format
        self.uniqueDates = []
        self.startTime = ''
        self.endTime = ''
        
        # For converted date strings QtDate used to form QtDateTime
        self.qtStartTime = QtCore.QTime()
        self.qtEndTime = QtCore.QTime()
        self.qtStartDate = QtCore.QDate()
        self.qtEndDate = QtCore.QDate()
              
        # Combine both into QtDateTime objects for comparing in script
        self.qtStartDateTime = QtCore.QDateTime()
        self.qtEndDateTime = QtCore.QDateTime()
        
        # Tool/menu bar clicks
        self.actionOpen.triggered.connect(self.openButton_clicked)
        self.actionQuit.triggered.connect(self.actionQuit_triggered)
        self.actionAbout.triggered.connect(self.actionAbout_triggered)
        self.actionTutorial.triggered.connect(self.actionTutorial_triggered)
        
        # Combo boxes
        self.startDateCombo.currentIndexChanged.connect(self.startDateCombo_changed)
        self.endDateCombo.currentIndexChanged.connect(self.endDateCombo_changed)
        
        # Time edit boxes
        self.startTimeEdit.editingFinished.connect(self.startTimeEdit_edited)
        self.startTimeEdit.editingFinished.connect(self.endTimeEdit_edited)
        
        # Check box doesn't need to connect to function, will check directly
        self.excelCheckBox.clicked.connect(self.excelCheckBox_clicked)
        
        # Push buttons
        self.openButton.clicked.connect(self.openButton_clicked)
        self.parseButton.clicked.connect(self.parseButton_clicked)


    def convertStartDate(self):
        """Converts start date string into a QDate object"""
        currStartDateString = self.startDateCombo.currentText() # get selection
        # convert string in combo box to QDate object
        tempQtDate = QtCore.QDate.fromString(currStartDateString, "MM/dd/yy")
        # convert and add 100 years (defaults 2 digit years to 1900's)
        self.qtStartDate = tempQtDate.addYears(100)
    
    def convertEndDate(self):
        """Converts end date string into a QDate object"""
        currEndDateString = self.endDateCombo.currentText()
        tempQtDate = QtCore.QDate.fromString(currEndDateString, "MM/dd/yy")
        self.qtEndDate = tempQtDate.addYears(100)
        
    def limitTimeEdit(self):
        qtMinLimit = QtCore.QTime.fromString(self.startTime, "hh:mm:ss")
        qtMaxLimit = QtCore.QTime.fromString(self.endTime, "hh:mm:ss")

        # If the dates are the same, end time can't be greater that start time
        # and the end time can't be less than the start time. Also limit selection
        # to what exists in the file.
        if self.startDateCombo.currentIndex() == self.endDateCombo.currentIndex():
            if self.startDateCombo.currentIndex() == 0:
                self.startTimeEdit.setMinimumTime(qtMinLimit)
                self.endTimeEdit.setMaximumTime(self.QT_MAX)
                self.endTimeEdit.setMinimumTime(self.startTimeEdit.time())
                self.startTimeEdit.setMaximumTime(self.endTimeEdit.time())
            
            elif self.startDateCombo.currentIndex() == len(self.uniqueDates)-1:
                self.startTimeEdit.setMinimumTime(self.QT_MIN)
                self.endTimeEdit.setMaximumTime(self.QT_MAX)
                self.endTimeEdit.setMinimumTime(self.startTimeEdit.time())
                self.startTimeEdit.setMaximumTime(self.endTimeEdit.time())
            
            else:
                self.startTimeEdit.setMinimumTime(self.QT_MIN)
                self.endTimeEdit.setMaximumTime(self.QT_MAX)
                self.endTimeEdit.setMinimumTime(self.startTimeEdit.time())
                self.startTimeEdit.setMaximumTime(self.endTimeEdit.time())
        
        else: # When dates are not equal: 
            # Set limits to the start time edit box
            if self.startDateCombo.currentIndex() == 0: # first day
                self.startTimeEdit.setMinimumTime(qtMinLimit)
                self.startTimeEdit.setMaximumTime(self.QT_MAX)
            
            # last day 
            elif self.startDateCombo.currentIndex() == len(self.uniqueDates)-1:
                self.startTimeEdit.setMinimumTime(qtMaxLimit)
                self.startTimeEdit.setMinimumTime(self.QT_MIN)
                 
            else: # any other mismatch day
                self.startTimeEdit.setMinimumTime(self.QT_MIN)
                self.startTimeEdit.setMaximumTime(self.QT_MAX)
                 
            # Set limits to the end time edit box
            if self.endDateCombo.currentIndex() == 0: # first day
                self.endTimeEdit.setMinimumTime(qtMinLimit)
                self.endTimeEdit.setMaximumTime(self.QT_MAX)
            
            # last day 
            elif self.endDateCombo.currentIndex() == len(self.uniqueDates)-1:
                self.endTimeEdit.setMaximumTime(qtMaxLimit)
                self.endTimeEdit.setMinimumTime(self.QT_MIN)
                 
            else: # any other mismatch day
                self.endTimeEdit.setMinimumTime(self.QT_MIN)
                self.endTimeEdit.setMaximumTime(self.QT_MAX)

        
        # Prevent start date from being later than end date
        if self.startDateCombo.currentIndex() > self.endDateCombo.currentIndex():
            self.startDateCombo.setCurrentIndex(self.endDateCombo.currentIndex())
            
        # Prevent end date from being earlier than start date
        if self.endDateCombo.currentIndex() < self.startDateCombo.currentIndex():
            self.endDateCombo.setCurrentIndex(self.startDateCombo.currentIndex())

        
    def updateQtDateTimes(self):
        """Update all current filter values to QDateTime objects and set min
        and max time values for start and end time edit boxes"""
        # Convert QDate and QTime obj into QDateTime objects for start dateTime
        self.qtStartTime = self.startTimeEdit.time() # update from gui edit box        
        startDateString = QtCore.QDate.toString(self.qtStartDate, "MM/dd/yyyy")
        startTimeString = QtCore.QTime.toString(self.qtStartTime, "hh:mm:ss")
        tempStartString = startDateString + ' ' + startTimeString
        self.qtStartDateTime = QtCore.QDateTime.fromString(tempStartString,"MM/dd/yyyy hh:mm:ss")

        # For end dateTime
        self.qtEndTime = self.endTimeEdit.time() # update from guit edit box
        endDateString = QtCore.QDate.toString(self.qtEndDate, "MM/dd/yyyy")
        endTimeString = QtCore.QTime.toString(self.qtEndTime, "hh:mm:ss")
        tempEndString = endDateString + ' ' + endTimeString
        self.qtEndDateTime = QtCore.QDateTime.fromString(tempEndString, "MM/dd/yyyy hh:mm:ss")  
    
    def startDateCombo_changed(self):
        self.convertStartDate()
        self.updateQtDateTimes()
        self.limitTimeEdit()
        
    def endDateCombo_changed(self):
        self.convertEndDate()
        self.updateQtDateTimes()
        self.limitTimeEdit()
        
    def startTimeEdit_edited(self):
        self.updateQtDateTimes()
        self.limitTimeEdit()
                
    def endTimeEdit_edited(self):
        self.updateQtDateTimes()
        self.limitTimeEdit()

    def getFileNameTuple(self):
        return self.fileNameTuple
                  
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
                
                # set the start and endDate combos to the first and last unique date
                self.startDateCombo.setCurrentIndex(1)
                self.endDateCombo.setCurrentIndex(len(self.uniqueDates)- 1)
                
                # Turn time strings into QTime obj and fill out time edits
                self.qtEndTime = QtCore.QTime.fromString(self.endTime, 'hh:mm:ss')
                self.qtStartTime = QtCore.QTime.fromString(self.startTime, 'hh:mm:ss')
                self.startTimeEdit.setTime(self.qtStartTime)
                self.endTimeEdit.setTime(self.qtEndTime)              
                               
                # Update all current QDateTime objects
                self.updateQtDateTimes()
                self.limitTimeEdit()
                  
                # Allow user to access filters and buttons
                self.enableButtons()

                # Change label to give user feedback
                self.changableLabel.setText(self.fullFileName + ' successfully loaded.')
            
            else:
                self.changableLabel.setText(self.fullFileName + ' is not a VitalView file')
    
    def enableButtons(self):
        self.actionOpen.setEnabled(1)
        self.openButton.setEnabled(1)
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
        
    def disableButtons(self):
        self.actionOpen.setEnabled(0)
        self.openButton.setEnabled(0)
        self.filterLabel.setEnabled(0)
        self.filterLine.setEnabled(0)
        self.startDateLabel.setEnabled(0)
        self.startDateCombo.setEnabled(0)
        self.startTimeLabel.setEnabled(0)
        self.startTimeEdit.setEnabled(0)
        self.endDateLabel.setEnabled(0)
        self.endDateCombo.setEnabled(0)
        self.endTimeEdit.setEnabled(0)
        self.excelCheckBox.setEnabled(0)
        self.parseButton.setEnabled(0)
    
    def lookForCsv2Excel(self):
        currWorkingDir = os.getcwd()
        csv2excelwbPath = os.path.join(currWorkingDir, 'tools', 'csv2excelwb.exe' )
        if os.path.exists(csv2excelwbPath):
            self.csv2excelwbPath = csv2excelwbPath
            return True
        else:
            msgBox = QtGui.QMessageBox()
            QtGui.QPlainTextEdit.LineWrapMode
            msgBox.setText("Cannot find csv2exelwb script/executable in the 'tools' folder.")
            msgBox.exec_()
            self.excelCheckBox.setChecked(False)

    def callCsv2Excel(self):
        if self.lookForCsv2Excel():
            subprocess.call([self.csv2excelwbPath, self.newFolderPath, self.fileNameNoExt])

    def excelCheckBox_clicked(self):
        self.lookForCsv2Excel()

    def parseButton_clicked(self):
        csv2excelwbPath = self.lookForCsv2Excel()
        rwparser = RwParser()
        savePath = self.dirName + "/" + self.fileNameNoExt
        if self.excelCheckBox.isChecked():
            # filter, calculate hourly, running streaks etc.
            
            rwparser.parseDistData(self.distCsvName, self.fileNameNoExt, 
                self.newFolderPath, self.qtStartDateTime, self.qtEndDateTime)
            self.changableLabel.setText("Parsing complete.")

            self.disableButtons()
            
            self.changableLabel.setText("Compiling excel worksheet. Please wait...")

            # Pop-up info box for user feedback
            msgBox = QtGui.QMessageBox()
            QtGui.QPlainTextEdit.LineWrapMode
            msgBox.setText("Compiling Excel workbook with csv output. Please wait. (Don't panic, the program may not respond for a few minutes.)")
            msgBox.exec_()

            # call csv2execlwb.py/.exe
            self.callCsv2Excel()

            # More feedback
            msgBox = QtGui.QMessageBox()
            msgBox.setText("Excel file complete! Results can be found in the " + savePath + " directory.")
            msgBox.exec_()
            
            # Export to Excel
            # This block is not currently used. Reserved for future. Hopefully,
            # I can figure out the segfault error and integrate the export
            # to excel portion back into the script.
            # rwparser.exportToExcel(self.fileNameNoExt, self.newFolderPath)
            
            self.changableLabel.setText("Finished compiling Excel workbook.")
            # Turn controls back on when done
            self.enableButtons()
        else: # pass onto
            rwparser.parseDistData(self.distCsvName, self.fileNameNoExt, 
                self.newFolderPath, self.qtStartDateTime, self.qtEndDateTime)
            self.changableLabel.setText("Parsing complete.")
            self.disableButtons()
            
            # Pop-up info box for user feedback
            msgBox = QtGui.QMessageBox()
            QtGui.QPlainTextEdit.LineWrapMode
            msgBox.setText("Parsing complete. Results can be found in the " + savePath + " directory ")
            msgBox.exec_()
            self.enableButtons()
          
if __name__ == '__main__':
        # File variables
        fullFilePath = ''
        app = QtGui.QApplication(sys.argv)
        MainApp = ParseWindow()
        MainApp.show()
        sys.exit(app.exec_())