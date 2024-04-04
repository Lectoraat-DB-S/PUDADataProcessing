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
COUNT_THE_AMOUNT = 240 + 2

# ERP excel sheet
excel_sheet1 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME1)
# Laser excel sheet
excel_sheet2 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME2)
# dataframe that will be filled we the results
data = pd.DataFrame(columns=[EXPORT_COLUMN_ERP_PROGRAMMANUMMER, EXPORT_COLUMN_ERP_TIJD, EXPORT_COLUMN_LASER_TIJD, EXPORT_COLUMN_LASER_PLAAT, EXPORT_COLUMN_LASER_PROGRAMMANUMMER])

print("are you even working")

# for index, row in excel_sheet2.iterrows():
#     if index == 232 or index == 233 or index == 234 or index == 235:
#         print("for real")
#         print(row[IMPORT_COLUMN_LASER_TIJD])
    # if isinstance(row[IMPORT_COLUMN_LASER_TIJD], pd._libs.tslibs.nattype.NaTType):
    #     print("you what mate")
    #     print(row[IMPORT_COLUMN_LASER_TIJD])

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

# variables used for finding all usable data is Laser sheet
lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
LastPlaatNr = 0
plaatNr = 0

# for index, row in excel_sheet2.iterrows():
#     if isinstance(row[IMPORT_COLUMN_LASER_TIJD], pd._libs.tslibs.nattype.NaTType):
#         print("you what mate")
#         print(row[IMPORT_COLUMN_LASER_TIJD])

# fill empty cells with '' because python can't handle NaN
excel_sheet2 = excel_sheet2.fillna('')

# variable used for debugging
counted = 0

# run through all laser data and match programnumber to time and to the plate in the laser sheet,
# also match the data from the ERP to it if there is an programnumber in the ERP data that matches a programnumber in the laser data
# for index, row in excel_sheet2.iterrows():
#     if index == 232 or index == 233 or index == 234 or index == 235:
#         print("for real")
#         print(row[IMPORT_COLUMN_LASER_TIJD])
#     if row[IMPORT_COLUMN_LASER_TIJD] == 'None' or (row[IMPORT_COLUMN_LASER_TIJD] == '' and (index == 232 or index == 233 or index == 234 or index == 235)):
#         print("you what mate")
#         print(row[IMPORT_COLUMN_LASER_TIJD])

# dataToBeRemoved = []

for index, row in excel_sheet2.iterrows():
    if row[IMPORT_COLUMN_LASER_TIJD] == '' and row[IMPORT_COLUMN_LASER_PLAAT] == '' and row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
        print("you son of a *****")
        # dataToBeRemoved.append(index)
        excel_sheet2.drop(index, axis='index', inplace=True)

# for i in dataToBeRemoved:
#     print(i)
#     excel_sheet2.drop(index=i, axis='index', inplace=True)

# excel_sheet2.drop([0,1], inplace=True)

excel_sheet2.to_excel(FILE_NAME)