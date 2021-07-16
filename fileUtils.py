import datetime
import sys
import os
import pandas as pd
from datetime import datetime as dt

def getPythonVersion():
    print(sys.version)
    print('--------')
    print(sys.version_info)
    print('---------')
    print(sys.version_info.msjor)
    if (sys.version_info.msjor >= 3):
        print('version >= 3')

#Find all files of a certain type with a certain prefix and store their date in a list.
#returns the list of datetimes
def findFilesOfTypeWithPrefix(prefix, fileType):
    inputFolderName = 'MDMSReport'
    #initialize list variable
    fileList = []
    #find the path to this program
    dir_path = os.path.join(findFilePath(),inputFolderName)

    #makes a list of all files in the directory where the program exists
    for root, dirs, files in os.walk(dir_path):
        #checks all files in the folder
        for file in files:
            #checks files for prefix and for file type
            if file.startswith(prefix) and file.endswith(fileType):           
                date = getDateFromFileName(file, prefix)
                fileList.append(date)
                
    # Returns a list of datetimes
    print('processing')
    return fileList

def findFilePath():
    #find the path to this program
    dir_path = os.path.dirname(os.path.realpath(__file__))
    return dir_path


def getDateFromFileName(fileName, prefix):
    #remove prefix from file name
    date_string = '/'.join(fileName.split(prefix)[1::])
    #remove file type
    date_string = date_string.strip( ".xlsx")
    
    #print(date_string)
    parsed_date = dt.strptime(date_string, "%d%b%Y")
    #print(parsed_date)
    return parsed_date

def getLatestDate(dates):
    newest = min(dates.iteritems(), key=lambda v: v if isinstance(v, datetime) else datetime.max)
    return newest

def createFolder(path = os.path.dirname(os.path.realpath(__file__))):
    if not os.path.isdir(path):
        print('No folder was found')

        # Make the folder
        os.mkdir(path)
        return True
    else:
        print('Folder Found')
        return False

def FolderCheck(folderName):
    # folderBool# is a Bool value that indicates whether or not the foler was created
    folderBool = createFolder(createPath(folderName))
    return folderBool


def isFolder(path):
    if not os.path.isdir(path):
        print('No folder was found')
        return False
    else:
        print('folder found')
        return True


def createPath(folder):
    path = os.path.join(findFilePath(), folder)
    return path

def findFiles(meterStatusList):
    if meterStatusList == []:
        print('No MDMS reports found!')
        input('Please place the MDMS reports into the proper folder and then hit Enter \'')
    else:
        print('Files Found.')


# If a folder wasn't found, wait for user to place files in folder
def filesInFolder(folderBool1, inputFolderName, meterStatusList):
    if folderBool1 == True:
        print('First Time Running or Folder not found')
        findFiles(meterStatusList)

        folderBool = isFolder(createPath(inputFolderName))
        if folderBool == False:
            print('Folder could not be created. Exiting program')
            exit()
    return meterStatusList
            

# These should be in excelUtils.py!!!!!!!!!!
# ... or should they?

def readExcelFile(fileName, sheetName='Sheet1', headerInfo = 3):
    if sys.version_info.major == 3 and sys.version_info.minor >=7:
        return pd.read_excel(fileName, sheet_name=sheetName, header=headerInfo)
    else:
        return pd.read_excel(fileName, sheet_name=sheetName, header=headerInfo, engine='openpyxl')

def writeExcelFile(fileName, dfs, folder = 'CompletedReport'):
    path1 = os.path.join(findFilePath(), folder)
    print (path1)
    writer = pd.ExcelWriter(fileName, engine='xlsxwriter')

    if fileName:
        # Loop through `dict` of dataframes
        for sheetname, df in dfs.items(): 
            # Send df to writer
            df.to_excel(writer, sheet_name=sheetname, index=False, header=True)
            # Pull worksheet object
            worksheet = writer.sheets[sheetname]
            # Get the shape of the dataframe
            (max_row, max_col) = df.shape
            
            # Become Table
            # ----------------------------------------------------------------------------
            column_settings = [{'header': column} for column in df.columns]
            worksheet.add_table(0, 0, max_row, max_col - 1, {'columns':column_settings})

            # Add Colors
            if sheetname == "Meter Status" or sheetname == "UP Meter Status":
                worksheet.conditional_format('D1:D500', {'type': '3_color_scale', 'min_color': '#5eff6c', 'max_color': '#ff7373'})
            elif sheetname == "Offline Meters" or sheetname == "UP Offline Meters":
                worksheet.conditional_format('F1:F500', {'type': '3_color_scale'})
            else:
                worksheet.conditional_format('D1:D500', {'type': '3_color_scale'})
            
            # Set the column width to the size of the largest content 
            # ----------------------------------------------------------------------------
            # Loop through all columns 
            for idx, col in enumerate(df):
                series = df[col]
                max_len = max((
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                    )) + 2  # adding a little extra space
                worksheet.set_column(idx, idx, max_len)  # set column width



        # Can I go home now?
        writer.save()
        