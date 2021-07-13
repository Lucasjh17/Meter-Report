import pandas as pd

def calculateOfflineMeters(df):
    #filter to show only sites
    sites = df.loc[df['Organization Level'] == 'Site']

    totalMeters = sites['Total Meters']
    totalCurrentMeters = sites['Total Current Meters']
    offlineMeters = totalMeters - totalCurrentMeters

    offlinePercent = (offlineMeters/totalMeters) * 100

    if not 'Offline Meters' in df.columns:
        df.insert(4,'Offline Meters', offlineMeters)
        df.insert(5,'Percent of Meters Offline', offlinePercent)

    return df

def sortByTextInAColumn(df, column, text):
    df = df.loc[df[column] == text]
    return df

def highlight_greaterthan(s, threshold, column):
    is_max = pd.Series(data=False, index=s.index)
    is_max[column] = s.loc[column] >= threshold
    return ['background-color: yellow' if is_max.any() else '' for v in is_max]


def getRowsBeforeString(df, column, text):

    df = df.loc[: df[(df[column] == text)].index[0], :]
    return df

def getRowsAfterString(df, column, text):

    df = df.loc[df[(df[column] == text)].index[0] :, :]
    return df
