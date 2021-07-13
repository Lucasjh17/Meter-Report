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
morningReport = 'Meter_Status_'
afternoonReport = 'Afternoon_Report_'

# Creates folder if not found
# folderBool#  indicates whether or not the folder was created
folderBool1 = fu.FolderCheck(inputFolderName)
folderBool2 = fu.FolderCheck(resultFolderName)


def menu():
    print('[1] Morning Report')
    print('[2] Afternoon Report')
    print('[3] Exit')
    n = input()

    if n == '1':
        # Collect variables to run morning report
        (currentReport, priorReport, currentReportDate, priorReportDate, fileName1) = ru.FindFilesForReport(morningReport)
        return (currentReport, priorReport, currentReportDate, priorReportDate, fileName1)
    elif n == '2':
        (currentReport, priorReport, currentReportDate, priorReportDate, fileName1) = ru.FindFilesForReport(afternoonReport)
        return (currentReport, priorReport, currentReportDate, priorReportDate, fileName1)
    elif n == '3':
        exit()
    else:
        menu()  
    
    

(currentReport, priorReport, currentReportDate, priorReportDate, fileName1) = menu()

# put each dataframe(excel sheet) into a dictionary
(dfDict,sites,offlineMeters,naughtyList) = ru.MorningReport(currentReport, priorReport, currentReportDate, priorReportDate)

# Name of the generated document
finalFileName = 'Completed_' + fileName1

# Generate xlsx file with each dataframe as its own sheet
fu.writeExcelFile(finalFileName, dfDict, resultFolderName)

show(sites,offlineMeters,naughtyList)

print('Finished')
