#     ___  ___  _________  ______                      _   _               _   _ _   _ _ _ _         
#    / _ \ |  \/  || ___ \ | ___ \                    | | (_)             | | | | | (_) (_) |        
#   / /_\ \| .  . || |_/ / | |_/ /___ _ __   ___  _ __| |_ _ _ __   __ _  | | | | |_ _| |_| |_ _   _ 
#   |  _  || |\/| ||  __/  |    // _ \ '_ \ / _ \| '__| __| | '_ \ / _` | | | | | __| | | | __| | | |
#   | | | || |  | || |     | |\ \  __/ |_) | (_) | |  | |_| | | | | (_| | | |_| | |_| | | | |_| |_| |
#   \_| |_/\_|  |_/\_|     \_| \_\___| .__/ \___/|_|   \__|_|_| |_|\__, |  \___/ \__|_|_|_|\__|\__, |
#                                   | |                            __/ |                       __/ |
#                                   |_|                           |___/                       |___/    
#
#   The goal of this program is to streamline the reporting process of AMP team members.
#   This is accomplished by processing excel files so that less manual processing is needed.
#
#   OS: only tested on Windows 10 64-bit
#   Dependencies: pandas, numpy, openpyxl, xlsxwriter
#                  pandasgui: pypiwin32
#   ---------------------------------------------------------------------------------
import fileUtils as fu
import reportUtils as ru
from pandasgui import show
# ----------------------------------------------------------------------------
inputFolderName = 'MDMSReport'
resultFolderName = 'CompletedReport'

# Creates folder if not found
# folderBool#  indicates whether or not the folder was created
folderBool1 = fu.FolderCheck(inputFolderName)
folderBool2 = fu.FolderCheck(resultFolderName)

# Collect variables to run morning report
(currentReport, priorReport, currentReportDate, priorReportDate, fileName1) = ru.FindFilesForReport()

# put each dataframe(excel sheet) into a dictionary
(dfDict,sites,offlineMeters,naughtyList) = ru.MorningReport(currentReport, priorReport, currentReportDate, priorReportDate)

# Name of the generated document
finalFileName = 'Completed_' + fileName1

# Generate xlsx file with each dataframe as its own sheet
fu.writeExcelFile(finalFileName, dfDict, resultFolderName)

show(sites,offlineMeters,naughtyList)

print('Finished')
