import os
import pandas as pd
import excelUtils as xu
import fileUtils as fu

# Declaring frequently used strings as variables
# ----------------------------------------------------------------------------
excelFile = '.xlsx'
meterStatus = 'Meter_Status_'
afternoonReport = 'Afternoon_Report_'
inputFolderName = 'MDMSReport'
resultFolderName = 'CompletedReport'
# ----------------------------------------------------------------------------

def FindFilesForReport():
    # Find meter status files
    meterStatusList = fu.findFilesOfTypeWithPrefix(meterStatus, excelFile)

    # Check to make sure the files are there
    while meterStatusList == []:
        input('Please place the MDMS Reports in the MDMSReport folder and then hit enter')
        meterStatusList = fu.findFilesOfTypeWithPrefix(meterStatus, excelFile)

    # Sort the list so the newest datetime is first
    meterStatusList.sort(reverse=True)

    # Get the two most recent datetimes
    currentReportDate = meterStatusList[0].strftime('%d%b%Y')
    priorReportDate = meterStatusList[1].strftime('%d%b%Y')

    # Recreate file name from datetime
    fileName1 = meterStatus + currentReportDate + excelFile
    fileName2 = meterStatus + priorReportDate + excelFile

    # Path to the files
    path1 = os.path.join(inputFolderName, fileName1)
    path2 = os.path.join(inputFolderName, fileName2)

    # Read excel file into a dataframe
    currentReport = fu.readExcelFile(path1)
    priorReport = fu.readExcelFile(path2)

    return (currentReport, priorReport, currentReportDate, priorReportDate, fileName1)

def MorningReport(currentReport, priorReport, currentReportDate, priorReportDate):
    # Only gets what comes before "  USARC"
    currentReport = xu.getRowsBeforeString(currentReport, 'Organization', '  USARC')
    priorReport = xu.getRowsBeforeString(priorReport, 'Organization', '  USARC')

    # Calculates offline meters in sheet one
    # Adds two new columns; Offline Meters, and Percent of Meters Offline
    currentReport = xu.calculateOfflineMeters(currentReport)
    priorReport = xu.calculateOfflineMeters(priorReport)

    # Filter the data to show only rows that have the Site tag in the Organiztion Level column
    sites = xu.sortByTextInAColumn(currentReport,'Organization Level', 'Site')
    sites2 = xu.sortByTextInAColumn(priorReport,'Organization Level', 'Site')

    # Remove Organization Level now that they're all sites
    sites = sites.drop(columns='Organization Level')

    # Get only the names and number of offline meters
    d1 = sites[['Organization','Offline Meters']].copy()
    d2 = sites2[['Organization','Offline Meters']].copy()

    # Rename the header so that the columns are labeled by date
    d1.columns = ['Organization',currentReportDate]
    d2.columns = ['Organization',priorReportDate]

    # Merge the two dataframes
    offlineMeters = pd.merge(d2, d1)

    # Get the difference between the two days.
    yesterday = offlineMeters[priorReportDate]
    today = offlineMeters[currentReportDate]
    meterDif = yesterday - today

    # Insert a new column with the new data
    offlineMeters.insert(3,'Difference Between Days', meterDif)

    # Make the Naughty List
    naughtyList = offlineMeters[['Organization', priorReportDate,currentReportDate,'Difference Between Days']].copy()
    naughtyList = naughtyList.loc[naughtyList[currentReportDate]>199]

    # Sort by absolute value
    offlineMeters = offlineMeters.iloc[offlineMeters['Difference Between Days'].abs().argsort()[::-1]]

    # Sort the Offline Meter values
    sites = sites.sort_values(['Offline Meters'], ascending = False,)

    # Add the dataframs to list to send them off to be exported
    dfDict = {"Meter Status": sites, "Offline Meters": offlineMeters, "Naughty List": naughtyList}
    return dfDict
