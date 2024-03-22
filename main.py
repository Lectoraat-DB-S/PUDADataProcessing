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
FILE_NAME = "result2.xlsx"

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
#sheet 2
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

# testing variables
COUNT_THE_AMOUNT = 9999 + 2

excel_sheet1 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME1)
# excel_sheet2 = pd.read_excel(file_location, sheet_name="Laser 3")
excel_sheet2 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME2)

# data = pd.DataFrame(columns=['ERP Programma Nummer', 'ERP TijdBonPerPlaat', 'Laser 3 GrossRunTime', 'Laser 3 Plaatnr', 'Laser 3 ProgrammaNaam'])
data = pd.DataFrame(columns=[EXPORT_COLUMN_ERP_PROGRAMMANUMMER, EXPORT_COLUMN_ERP_TIJD, EXPORT_COLUMN_LASER_TIJD, EXPORT_COLUMN_LASER_PLAAT, EXPORT_COLUMN_LASER_PROGRAMMANUMMER])

index_new_data = 0
last_programmanummer = 0

for index, row in excel_sheet1.iterrows():
    if row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] != last_programmanummer:
        last_programmanummer = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]
        data.loc[index_new_data] = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER], 0, 0, 0, 0
        index_new_data = index_new_data + 1

for index, row in excel_sheet1.iterrows():
    placementIndex1 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]].tolist()
    data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] + row[IMPORT_COLUMN_ERP_TIJD]

lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
FirstGrossRunTime = 0
LastPlaatNr = 0
plaatNr = 0

excel_sheet2 = excel_sheet2.fillna('')

counted = 0

for index, row in excel_sheet2.iterrows():
    if DEBUGGING:
        print("start")
    if lastProgramNumber == 0:
        lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
    if DEBUGGING:
        print("check for number")
    if lastProgramNumber == row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] or row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
        if DEBUGGING:
            print("matched number")
            print("check for plaatnr")
        if row[IMPORT_COLUMN_LASER_PLAAT] == LastPlaatNr or row[IMPORT_COLUMN_LASER_PLAAT] == '':
            if DEBUGGING:
                print("matched plaatnr")
                print("check for grossruntime")
            if row[IMPORT_COLUMN_LASER_TIJD] != '':
                if DEBUGGING:
                    print("matched grossruntime")
                LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]
        else:
            if DEBUGGING:
                print("not a matched plaatnr")
            plaatNr = LastPlaatNr
            LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            programmaNummer = lastProgramNumber
            GrossRunTime = LastGrossRunTime
            LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]
            placementIndex2 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == lastProgramNumber].tolist()
            if not placementIndex2:
                if DEBUGGING:
                    print('list is empty')
                test1 = ''
                test2 = ''
            else:
                test1 = data.iat[placementIndex2[0], 0]
                test2 = data.iat[placementIndex2[0], 1]
            if 1:
                test3 = GrossRunTime
                test4 = plaatNr
                test5 = programmaNummer
                new_row = {EXPORT_COLUMN_ERP_PROGRAMMANUMMER: test1, EXPORT_COLUMN_ERP_TIJD: test2, EXPORT_COLUMN_LASER_TIJD: test3, EXPORT_COLUMN_LASER_PLAAT: test4, EXPORT_COLUMN_LASER_PROGRAMMANUMMER: test5}
                data.loc[len(data)+1] = new_row
    else:
        if DEBUGGING:
            print("not a matched number")
        lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
        LastPlaatNr = 0

    if DEBUGGING:
        if counted == COUNT_THE_AMOUNT:
            break
        else:
            counted = counted + 1




data.to_excel(FILE_NAME)

result = pd.read_excel(FILE_NAME)
if DEBUGGING:
    print(result)
