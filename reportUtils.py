import os
import pandas as pd
import excelUtils as xu
import fileUtils as fu

# Declaring frequently used strings as variables
# ----------------------------------------------------------------------------
excelFile = '.xlsx'

inputFolderName = 'MDMSReport'
resultFolderName = 'CompletedReport'
# ----------------------------------------------------------------------------

def FindFilesForReport(fileName):
    # Find meter status files
    meterStatusList = fu.findFilesOfTypeWithPrefix(fileName, excelFile)

    # Check to make sure the files are there
    while meterStatusList == []:
        input('Please place the MDMS Reports in the MDMSReport folder and then hit enter')
        meterStatusList = fu.findFilesOfTypeWithPrefix(fileName, excelFile)

    # Sort the list so the newest datetime is first
    meterStatusList.sort(reverse=True)

    # Get the two most recent datetimes
    currentReportDate = meterStatusList[0].strftime('%d%b%Y')
    priorReportDate = meterStatusList[1].strftime('%d%b%Y')

    # Recreate file name from datetime
    fileName1 = fileName + currentReportDate + excelFile
    fileName2 = fileName + priorReportDate + excelFile

    # Path to the files
    path1 = os.path.join(inputFolderName, fileName1)
    path2 = os.path.join(inputFolderName, fileName2)

    # Read excel file into a dataframe
    currentReport = fu.readExcelFile(path1)
    priorReport = fu.readExcelFile(path2)

    return (currentReport, priorReport, currentReportDate, priorReportDate, fileName1)

def MorningReport(currentReport, priorReport, currentReportDate, priorReportDate):

    # Account Managers
    # is this bad programming or is this what weak typing is for?
    accountManagers = os.path.join(fu.findFilePath(),'Account_Managers.xlsx')
    accountManagers = fu.readExcelFile(accountManagers, headerInfo=0)
    accountManagers = accountManagers[['Organization','Account Manager','Back-Up']].copy()
    
    accountUPManagers = os.path.join(fu.findFilePath(),'UP_Account_Managers.xlsx')
    accountUPManagers = fu.readExcelFile(accountUPManagers, headerInfo=0)
    accountUPManagers = accountUPManagers[['Organization','Account Manager','Back-Up']].copy()

    # Only gets what comes before "  USARC", which should be the IMCOM sites
    currentIMCOMSites = xu.getRowsBeforeString(currentReport, 'Organization', '  USARC')
    priorIMCOMsites = xu.getRowsBeforeString(priorReport, 'Organization', '  USARC')

    # Get the AMC sites
    currentAMCSites = xu.getRowsBeforeString(currentReport, 'Organization Level', 'UP Meters')
    priorAMCSites = xu.getRowsBeforeString(priorReport, 'Organization Level', 'UP Meters')

    # Get the Utility Provider sites
    currentUPSites = xu.getRowsAfterString(currentReport, 'Organization Level', 'UP Meters')
    priorUPSites = xu.getRowsAfterString(currentReport, 'Organization Level', 'UP Meters')
 

    currentAMCSites = xu.getRowsAfterString(currentAMCSites, 'Organization', '  AMC')
    priorAMCSites = xu.getRowsAfterString(priorAMCSites, 'Organization', '  AMC')
    currentAMCSites = xu.sortByTextInAColumn(currentUPSites,'Organization Level', 'Site')
    
    currentSites = [currentAMCSites, currentIMCOMSites]
    priorSites = [priorAMCSites, priorIMCOMsites]

    currentReport = pd.concat(currentSites)
    priorReport = pd.concat(priorSites)



    # Calculates offline meters in sheet one
    # Adds two new columns; Offline Meters, and Percent of Meters Offline
    currentReport = xu.calculateOfflineMeters(currentReport)
    priorReport = xu.calculateOfflineMeters(priorReport)

    currentUPSites = xu.calculateOfflineMeters(currentUPSites)
    priorUPSites = xu.calculateOfflineMeters(priorUPSites)  

    # Filter the data to show only rows that have the Site tag in the Organiztion Level column
    sites = xu.sortByTextInAColumn(currentReport,'Organization Level', 'Site')
    sites2 = xu.sortByTextInAColumn(priorReport,'Organization Level', 'Site')

    currentUPSites = xu.sortByTextInAColumn(currentUPSites,'Organization Level', 'Site')
    priorUPSites = xu.sortByTextInAColumn(priorUPSites,'Organization Level', 'Site')


    # Remove Organization Level now that they're all sites
    sites = sites.drop(columns='Organization Level')
    currentUPSites = currentUPSites.drop(columns='Organization Level')

    # Get only the names and number of offline meters
    d1 = sites[['Organization','Offline Meters']].copy()
    d2 = sites2[['Organization','Offline Meters']].copy()

    uC = currentUPSites[['Organization','Offline Meters']].copy()
    uP = priorUPSites[['Organization','Offline Meters']].copy()


    # Rename the header so that the columns are labeled by date
    d1.columns = ['Organization',currentReportDate]
    d2.columns = ['Organization',priorReportDate]

    uC.columns = ['Organization',currentReportDate]
    uP.columns = ['Organization',priorReportDate]
    

    # Merge the two dataframes
    offlineMeters = pd.merge(d2, d1)
    offlineMeters = pd.merge(accountManagers, offlineMeters, left_on='Organization', right_on='Organization')
    
    offlineUPMeters = pd.merge(uP, uC)
    offlineUPMeters = pd.merge(accountUPManagers, offlineUPMeters, left_on='Organization', right_on='Organization')

    def getMeterDif(df):
        # Get the difference between the two days.
        yesterday = df[priorReportDate]
        today = df[currentReportDate]
        meterDif = yesterday - today
        return meterDif

    meterDif = getMeterDif(offlineMeters)
    meterUPDif = getMeterDif(offlineUPMeters)

    # Insert a new column with the new data
    offlineMeters.insert(5,'Difference Between Days', meterDif)

    offlineUPMeters.insert(5,'Difference Between Days', meterUPDif)

    # Make the Naughty List
    naughtyList = offlineMeters[['Organization', priorReportDate,currentReportDate,'Difference Between Days']].copy()
    naughtyList = naughtyList.loc[naughtyList[currentReportDate]>199]

    # Sort by absolute value
    offlineMeters = offlineMeters.iloc[offlineMeters['Difference Between Days'].abs().argsort()[::-1]]
    offlineUPMeters = offlineUPMeters.iloc[offlineUPMeters['Difference Between Days'].abs().argsort()[::-1]]


    # Sort the Offline Meter values
    sites = sites.sort_values(['Offline Meters'], ascending = False,)
    currentUPSites = currentUPSites.sort_values(['Offline Meters'], ascending = False,)

    # Add the dataframs to list to send them off to be exported
    dfDict = {"Meter Status": sites, "Offline Meters": offlineMeters, "Naughty List": naughtyList, "UP Meter Status": currentUPSites, "UP Offline Meters": offlineUPMeters}

    

    return (dfDict,sites,offlineMeters,naughtyList)
    
