# prw-gui
parse-running-wheel gui version for windows

Requires:
- Python 2.7, and pywin32 (http://sourceforge.net/projects/pywin32/). Make sure your pywin32 matches your python (e.g. 64bit python needs 64bit pywin32)
- Use 'pip install <package_name> for PySide, openpyxl, and Pyinstaller

Compiling instructions for Windows
- Install all dependencies, pyinstaller and pywin32 
- Run "$ pyinstaller -F --noconsole csv2excelwb.py"
- Run "$ pyinstaller -F --noconsole prw-gui_main.py"
- Create a new directory called 'tools' and place the compiled 'csv2exelwb.exe' into the folder
- Leave the prw-gui_main.exe in the parent folder
- Run prw-gui_main.exe to parse data!

Note to run source from commandline or terminal or compile for another OS
- Adjust the callCsvToExcel() function lines 309-310 of prw-gui_main.py to call 'csv2excelwb.py' instead of 'csv2exelwb.exe'


#########################################
#### Parse Running Wheel GUI Version ####
#########################################

By Prech Uapinyoying
prw-gui - Version 0.1
https://github.com/puapinyoying

############
# Tutorial #
############

This program parses mouse running wheel data that has been output into an ASCII 
(.asc) file directly obtained from the Vitalview Software. The input file should
have two headers. The first will be the the over all file header that says 
'Experiment Logfile:' followed by the date. Second, the data header below in 
comma delimited format (csv).

The program will output a csv file for each of the following:

        1) Mice data - unmanipulated summary statistics from ASCII file

        2) Raw data - unmanipulated data from ASCII file

        3) Distance calculated - converted turns data into meters

        4) Time Filtered - option to set start and end date. If unchanged, all
           data will be used. This sheet also includes a few summary statistics
           calculated at the bottom rows.

        5) Sum Hourly - minute data is summed into hours per row

        6) Cumulative - data from each hour is compounded onto the previous

        7) Running Streaks - how many consecutive minutes each animal runs per
           session. This includes all running sessions for that particular
           animal.

Instructions: 

        1) To start click "Open File" and select a VitalView (.asc) data file.

        2) Once loaded, if you would like to filter the data by start and end
           date/time use the drop-down boxes and time edit boxes (time format in
           hh:mm:ss)

        3) The program will automatically output csv files of all the
           calculations listed above. Option to export the csv files into a
           single Excel workbook.

        4) Click "Parse Data" to start."""