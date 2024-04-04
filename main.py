import pandas as pd
import numpy as np
from openpyxl import load_workbook

# setup constants
# put a 1 if debugging should be enabled
DEBUGGING = 0
MACHINE_NUMBER = 4
if MACHINE_NUMBER != 3 and MACHINE_NUMBER != 4:
    print("wrong machine number used")
    print("acceptable machine numbers are 3 and 4")
    print("shutting down, toedeloe")
    exit()

# declare constants
FILE_LOCATION = "Data.xlsx"
if MACHINE_NUMBER == 3:
    FILE_NAME = "resultLaser3V2.xlsx"
elif MACHINE_NUMBER == 4:
    FILE_NAME = "resultLaser4V2.xlsx"

# sheet names of the imported file
EXCEL_SHEET_NAME1 = "ERP"
if MACHINE_NUMBER == 3:
    EXCEL_SHEET_NAME2 = "Laser 3"
elif MACHINE_NUMBER == 4:
    EXCEL_SHEET_NAME2 = "Laser4"

# names of the columns from the imported file
# sheet 1
IMPORT_COLUMN_ERP_PROGRAMMANUMMER = 'Programmanummer'
IMPORT_COLUMN_ERP_TIJD = 'TijdBonPerPlaat'
# sheet 2
IMPORT_COLUMN_LASER_TIJD = 'GrossRunTime'
IMPORT_COLUMN_LASER_PLAAT = 'PlaatNr'
IMPORT_COLUMN_LASER_PROGRAMMANUMMER = 'ProgrammaNaam'

# names of the exported columns
EXPORT_COLUMN_ERP_PROGRAMMANUMMER = 'ERP Programma Nummer'
EXPORT_COLUMN_ERP_TIJD = 'ERP TijdBonPerPlaat'
if MACHINE_NUMBER == 3:
    EXPORT_COLUMN_LASER_TIJD = 'Laser 3 GrossRunTime'
    EXPORT_COLUMN_LASER_PLAAT = 'Laser 3 Plaatnr'
    EXPORT_COLUMN_LASER_PROGRAMMANUMMER = 'Laser 3 ProgrammaNaam'
elif MACHINE_NUMBER == 4:
    EXPORT_COLUMN_LASER_TIJD = 'Laser 4 GrossRunTime'
    EXPORT_COLUMN_LASER_PLAAT = 'Laser 4 Plaatnr'
    EXPORT_COLUMN_LASER_PROGRAMMANUMMER = 'Laser 4 ProgrammaNaam'

# testing constant
COUNT_THE_AMOUNT = 260 + 2

# ERP excel sheet
excel_sheet1 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME1)
# Laser excel sheet
excel_sheet2 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME2)
# dataframe that will be filled we the results
data = pd.DataFrame(columns=[EXPORT_COLUMN_ERP_PROGRAMMANUMMER, EXPORT_COLUMN_ERP_TIJD, EXPORT_COLUMN_LASER_TIJD, EXPORT_COLUMN_LASER_PLAAT, EXPORT_COLUMN_LASER_PROGRAMMANUMMER])

# indexing and checking for removing all duplicate programnumbers
index_new_data = 0
last_programmanummer = 0

# Find all programnumbers that exist in the ERP sheet
for index, row in excel_sheet1.iterrows():
    if row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] != last_programmanummer:
        last_programmanummer = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]
        data.loc[index_new_data] = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER], 0, 0, 0, 0
        index_new_data = index_new_data + 1

# add total time found in ERP sheet to all the programnumbers found in ERP sheet
for index, row in excel_sheet1.iterrows():
    # index for knowing where to find the data with the correct programnumber in the ERP sheet
    placementIndex1 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]].tolist()
    data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] + row[IMPORT_COLUMN_ERP_TIJD]

# for indexx, row in excel_sheet2.iterrows():
#     if isinstance(row[IMPORT_COLUMN_LASER_TIJD], pd._libs.tslibs.nattype.NaTType):
#         print(row[IMPORT_COLUMN_LASER_TIJD])
#         print(row[IMPORT_COLUMN_LASER_PLAAT])
#         excel_sheet2.drop(indexx, axis='index')

# excel_sheet2.to_excel("testData.xlsx")

# variables used for finding all usable data is Laser sheet
lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
LastPlaatNr = 0
plaatNr = 0

# fill empty cells with '' because python can't handle NaN
excel_sheet2 = excel_sheet2.fillna('')
# excel_sheet2.to_excel("testDataV2.xlsx")

for index, row in excel_sheet2.iterrows():
    if row[IMPORT_COLUMN_LASER_TIJD] == '' and row[IMPORT_COLUMN_LASER_PLAAT] == '' and row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
        excel_sheet2.drop(index, axis='index', inplace=True)

# excel_sheet2.to_excel("tempData.xlsx")

# variable used for debugging
counted = 0

# run through all laser data and match programnumber to time and to the plate in the laser sheet,
# also match the data from the ERP to it if there is an programnumber in the ERP data that matches a programnumber in the laser data
for index, row in excel_sheet2.iterrows():
    if DEBUGGING:
        print("start")
    # for the first run safe a number to prevent skipping
    if lastProgramNumber == 0:
        lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
    if DEBUGGING:
        print("check for number")
        print(lastProgramNumber)
        print(LastPlaatNr)
        print(counted)
        print(row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER])
    # check if the same programnumber is found (or if it was an empty cell, which impies the same program number)
    if lastProgramNumber == row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] or row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
        if DEBUGGING:
            print("matched number")
            print("check for plaatnr")
        # check if the same plate number was found (or if it was an empty cell, which impies the same plate number)
        if row[IMPORT_COLUMN_LASER_PLAAT] == LastPlaatNr or row[IMPORT_COLUMN_LASER_PLAAT] == '':
            if DEBUGGING:
                print("matched plaatnr")
                print("check for grossruntime")
            # check if the cell was not empty, because then the time was changed or the same and needs to be safed
            if row[IMPORT_COLUMN_LASER_TIJD] != '':
                if DEBUGGING:
                    print("matched grossruntime")
                LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]
        # if not, then it was the end of the plate and we should safe the data en continue to the next plate
        else:
            if DEBUGGING:
                print("not a matched plaatnr")
            # for 2 variables, safe them in a different variable and then get the next data
            # for 1 variable, doesn't matter, it always contains the latest data
            plaatNr = LastPlaatNr
            LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            programmaNummer = lastProgramNumber
            GrossRunTime = LastGrossRunTime
            LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]

            # index for knowing where to find the data with the same programnumber in the ERP sheet
            placementIndex2 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == lastProgramNumber].tolist()

            # in case the same programnumber from the laser sheet could not be found in the ERP sheet, leave cells empty
            if not placementIndex2:
                if DEBUGGING:
                    print('list is empty')
                ERPProgrammanummer = ''
                ERPTijd = ''
            # in case the same programnumber was found, fill the cells with the correct data
            else:
                ERPProgrammanummer = data.iat[placementIndex2[0], 0]
                ERPTijd = data.iat[placementIndex2[0], 1]

            # create the new row that will have to be added to the dataframe
            new_row = {EXPORT_COLUMN_ERP_PROGRAMMANUMMER: ERPProgrammanummer, EXPORT_COLUMN_ERP_TIJD: ERPTijd, EXPORT_COLUMN_LASER_TIJD: GrossRunTime, EXPORT_COLUMN_LASER_PLAAT: plaatNr, EXPORT_COLUMN_LASER_PROGRAMMANUMMER: programmaNummer}
            # add the new row to the dataframe
            data.loc[len(data)+1] = new_row
    # if not, it means we switched programnumber and by definition started from plat 0 (this might be incorrect and needs to be checked with the data)
    # elif LastGrossRunTime == 'None':
    #     print("this is fucking bullshit")

    # elif LastGrossRunTime != '' and LastGrossRunTime != 'None':
    else:
        if DEBUGGING:
            print("not a matched program number")
            print(LastGrossRunTime)
        if LastGrossRunTime == '':
            lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
            if row[IMPORT_COLUMN_LASER_PLAAT] != '':
                LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            else:
                LastPlaatNr = 0
            LastGrossRunTime = 0
        elif LastGrossRunTime != '':
            # print("what")
            # print(lastProgramNumber)
            # print(LastPlaatNr)
            # print(LastGrossRunTime)
            plaatNr = LastPlaatNr
            if row[IMPORT_COLUMN_LASER_PLAAT] != '':
                LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            else:
                LastPlaatNr = 0
            programmaNummer = lastProgramNumber
            lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
            GrossRunTime = LastGrossRunTime
            LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]

            # index for knowing where to find the data with the same programnumber in the ERP sheet
            placementIndex2 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == programmaNummer].tolist()

            # in case the same programnumber from the laser sheet could not be found in the ERP sheet, leave cells empty
            if not placementIndex2:
                if DEBUGGING:
                    print('list is empty')
                ERPProgrammanummer = ''
                ERPTijd = ''
            # in case the same programnumber was found, fill the cells with the correct data
            else:
                ERPProgrammanummer = data.iat[placementIndex2[0], 0]
                ERPTijd = data.iat[placementIndex2[0], 1]

            # create the new row that will have to be added to the dataframe
            new_row = {EXPORT_COLUMN_ERP_PROGRAMMANUMMER: ERPProgrammanummer, EXPORT_COLUMN_ERP_TIJD: ERPTijd, EXPORT_COLUMN_LASER_TIJD: GrossRunTime, EXPORT_COLUMN_LASER_PLAAT: plaatNr, EXPORT_COLUMN_LASER_PROGRAMMANUMMER: programmaNummer}
            # add the new row to the dataframe
            data.loc[len(data)+1] = new_row

    # debugging functions which enables only part of the excel sheet from being checked
    if DEBUGGING:
        if counted == COUNT_THE_AMOUNT:
            break
        else:
            counted = counted + 1

for index, row in data.iterrows():
    if row[EXPORT_COLUMN_LASER_PLAAT] == 0 and row[EXPORT_COLUMN_LASER_PROGRAMMANUMMER] == 0 and row[EXPORT_COLUMN_LASER_TIJD] == 0:
        data.drop(index, axis='index', inplace=True)

# write dataframe with results to a new excel file
data.to_excel(FILE_NAME)

# read the excel file with the results
if DEBUGGING:
    result = pd.read_excel(FILE_NAME)
    print(result)
