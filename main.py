import math

from template import template
from template import templateByBU
from ldosImport import ldosSummary
from ldosImport import summaryItem
from TableBuilder import TableItem
import os
import glob
import pandas as pd
import time


def main():
    startTime = time.time()

    #os.chdir('LDOS_Files')
    os.chdir('Holding')
    FileList = glob.glob('*.xlsx')
    for file in FileList:
        #Get templates to connect to mapping
        templates = getTemplates('../Files/WPA Consolidated BoMs v2.xlsx')

        #Get data mapping to templates
        mData = getMappedData(file)
        newTemplates = mergeMapAndTemplate(templates, mData)

        #Print new files
        fn = str(file).replace('Completed', 'Merged')
        printTemplates(newTemplates, fn)

    print(f'\nCompleted in {round(time.time()-startTime, 2)}s')

def mergeMapAndTemplate(templates, mData):
    print('Merging template and mapping data...')
    # Combine data mapping to templates
    custTemplates = templates
    for template in custTemplates:
        # Test to see if mData has any data for the BU
        sumItems = list(filter(lambda x : x.BU == template.BU, mData))

        if len(sumItems) > 0:
            sumItems = sumItems[0].sItems
        else :
            continue

        for t in template.templates :
            sumItem = list(filter(lambda x : x.name == t.name, sumItems))
            if len(sumItem) > 0:
                sumItem = sumItem[0]
                t.dates = sort(sumItem.dates)
                multiplier = sumItem.qty
            else:
                multiplier = 0

            for i in t.items :
                i.Qty *= multiplier
    return custTemplates

def sort(unsortedList):
    myKeys = list(unsortedList.keys())
    myKeys.sort()
    sorted_dict = {i : unsortedList[i] for i in myKeys}

    return sorted_dict


def printTemplates(templates, fileName):
    print('Exporting to excel...')
    tbl = TableItem(templates)

    tbl.printRows(fileName)


def getTemplates(path) :
    print('\nRefreshing Templates...')
    path = path

    #Get Worksheet names
    tabs = pd.ExcelFile(path, engine='openpyxl').sheet_names
    templates = []

    #Open each sheet and build template
    for sheet in tabs:
        t = templateByBU()

        df = pd.read_excel(path, sheet, engine='openpyxl')

        t.BU = sheet
        if not df.empty :
            t.templates = excelTemplateImport(df, sheet)

        templates.append(t)

    return templates


def excelMappedImport(df, sheet) :
    uniqItems = df['Replacement'].unique()
    uniqDates = df['LDOS Year'].unique()

    uniqDates = [d for d in uniqDates if str(d) != 'nan']

    sItems = []

    for item in uniqItems:
        replaceDF = df[df['Replacement'] == item]
        count = df[df['Replacement'] == item].shape[0]
        if count > 1:
            dates = {}
            dateCount = 0

            for date in uniqDates:
                dateCount = replaceDF[replaceDF['LDOS Year'] == date].shape[0]
                dates[date] = dateCount

            sItems.append(summaryItem(item, count, dates))

    return sItems


def getMappedData(path):
    path = path
    print(f'Opening {path}...')

    # Get Worksheet names
    tabs = pd.ExcelFile(path, engine='openpyxl').sheet_names
    sumItems = []

    allSheets = pd.read_excel(path, sheet_name=None, engine='openpyxl')

    # Open each sheet and build template
    for sheet in tabs:
        print(f'     Mapping {sheet}...')
        s = ldosSummary()

        df = allSheets[sheet]
        #df = pd.read_excel(path, sheet, engine='openpyxl')

        s.BU = sheet

        if not df.empty:
            if 'Replacement' in df.columns:
                s.sItems = excelMappedImport(df, sheet)

        sumItems.append(s)

    return sumItems


def excelTemplateImport(df, worksheet):

    templates = []

    #Find first header row
    startRow = df.index[df['Unnamed: 0'] == 'Line Number'].values.min()

    #Drop rows above the header
    df.drop(range(0, startRow), inplace=True)

    #Set first row as header
    cols = df.iloc[0]
    df = df[1:]
    df.columns = cols

    #Find all instances of BOMS
    templateList = df.index[df['Line Number'] == 'WPA_Name'].values
    templateList = [x - (startRow + 1) for x in templateList]
    templateList.append(len(df.index))

    #Assign data to object
    prevSection = 0

    for secEnd in templateList:
        t = template()
        t.BU = worksheet
        haveName = False
        for r in range(prevSection, secEnd):
            if not haveName:
                t.name = df.iloc[r]['Part Number']
                haveName = True
                continue
            row = df.iloc[r]
            #if isfloat(row['Line Number']) and not math.isnan(float(row['Line Number'])):
            if len(str(row['Line Number'])) < 12 and "." in str(row['Line Number']):
                t.appendItem(row['Line Number'], row['Part Number'], row['Description'], float(row['Unit List Price']),
                             int(row['Qty']), row['Disc(%)'])
        t.calcSubtotal()
        if len(t.items) > 0:
            templates.append(t)
        prevSection = secEnd

    return templates


def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

if __name__ == '__main__':
    main()


