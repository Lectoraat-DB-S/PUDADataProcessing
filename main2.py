import pandas as pd
import numpy as np
from openpyxl import load_workbook

# setup constants
# put a 1 if debugging should be enabled
DEBUGGING = 0

# declare constants
FILE_LOCATION = "SuplaconQ1.xlsx"
FILE_NAME = "SuplaconQ1V2.xlsx"

# sheet names of the imported file
EXCEL_SHEET_NAME1 = "Alle platen"

# names of the columns from the imported file
# sheet 1 ERP
IMPORT_COLUMN_PROGRAMMANUMMER = 'Programmanummer'
IMPORT_COLUMN_AANTALPLATEN = 'Aantal platen per programma'
IMPORT_COLUMN_MATERIAALCODE = 'Materiaalcode'
IMPORT_COLUMN_STUKS = 'TotaalAantalStuksPerPlaat'
IMPORT_COLUMN_TIJDPERPLAAT = 'TotaleTijdPerPlaat'
IMPORT_COLUMN_TOTALEINSTEKEN = 'InstekenTotaleAantal'
IMPORT_COLUMN_TOTALESNIJLENGTE = 'TotaleSnijlengtePerPlaat'
IMPORT_COLUMN_TIJD_METER = 'Tijd/meter'
IMPORT_COLUMN_GROSS_TIME = 'Gross time '
IMPORT_COLUMN_NETTO_TIME = 'Netto time'
IMPORT_COLUMN_PROCESSING_TIME = 'Processing time'
IMPORT_COLUMN_DIFFERENCE_VCNC = 'Verschil vc-nc'
IMPORT_COLUMN_PLAATNUMMER = 'Plaatnummer'
IMPORT_COLUMN_LASER = 'Laser'

# names of the exported columns
EXPORT_COLUMN_PROGRAMMANUMMER = 'Programmanummer'
EXPORT_COLUMN_AANTALPLATEN = 'Aantal platen per programma'
EXPORT_COLUMN_MATERIAALCODE = 'Materiaalcode'
EXPORT_COLUMN_STUKS = 'TotaalAantalStuksPerPlaat'
EXPORT_COLUMN_TIJDPERPLAAT = 'TotaleTijdPerPlaat'
EXPORT_COLUMN_TOTALEINSTEKEN = 'InstekenTotaleAantal'
EXPORT_COLUMN_TOTALESNIJLENGTE = 'TotaleSnijlengtePerPlaat'
EXPORT_COLUMN_TIJD_METER = 'Tijd/meter'
EXPORT_COLUMN_GROSS_TIME = 'Gross time'
EXPORT_COLUMN_NETTO_TIME = 'Netto time'
EXPORT_COLUMN_PROCESSING_TIME = 'Processing time'
EXPORT_COLUMN_DIFFERENCE_VCNC = 'Verschil vc-nc'
EXPORT_COLUMN_PLAATNUMMER = 'Plaatnummer'
EXPORT_COLUMN_LASER = 'Laser'

# testing constant
COUNT_THE_AMOUNT = 260 + 2

# ERP excel sheet
excel_sheet1 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME1)
# # WICAM excel sheet
# excel_sheet2 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME2)
# # Laser excel sheet
# excel_sheet3 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME3)
# dataframe that will be filled we the results
data = pd.DataFrame(columns=[EXPORT_COLUMN_PROGRAMMANUMMER,
                             EXPORT_COLUMN_AANTALPLATEN,
                             EXPORT_COLUMN_MATERIAALCODE,
                             EXPORT_COLUMN_STUKS,
                             EXPORT_COLUMN_TIJDPERPLAAT,
                             EXPORT_COLUMN_TOTALEINSTEKEN,
                             EXPORT_COLUMN_TOTALESNIJLENGTE,
                             EXPORT_COLUMN_TIJD_METER,
                             EXPORT_COLUMN_GROSS_TIME,
                             EXPORT_COLUMN_NETTO_TIME,
                             EXPORT_COLUMN_PROCESSING_TIME,
                             EXPORT_COLUMN_DIFFERENCE_VCNC,
                             EXPORT_COLUMN_PLAATNUMMER,
                             EXPORT_COLUMN_LASER])

# excel_sheet1[IMPORT_COLUMN_ERP_ORDER_BON] = excel_sheet1[IMPORT_COLUMN_ERP_ORDER_BON].str.replace('.', '-')

# # indexing and checking for removing all duplicate programnumbers
index_new_data = 0
last_programmanummer = 0

# # Find all programnumbers that exist in the ERP sheet
# for index, row in excel_sheet1.iterrows():
#     if row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] != last_programmanummer:
#         last_programmanummer = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]
#         data.loc[index_new_data] = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER], 0, 0, 0, 0, 0, 0, 0, 0
#         index_new_data = index_new_data + 1
        
for index, row in excel_sheet1.iterrows():
    # data.loc[index] = EXPORT_COLUMN_PROGRAMMANUMMER,
    #                     EXPORT_COLUMN_AANTALPLATEN,
    #                     EXPORT_COLUMN_MATERIAALCODE,
    #                     EXPORT_COLUMN_STUKS,
    #                     EXPORT_COLUMN_TIJDPERPLAAT,
    #                     EXPORT_COLUMN_TOTALEINSTEKEN,
    #                     EXPORT_COLUMN_TOTALESNIJLENGTE,
    #                     EXPORT_COLUMN_TIJD_METER,
    #                     EXPORT_COLUMN_GROSS_TIME,
    #                     EXPORT_COLUMN_NETTO_TIME,
    #                     EXPORT_COLUMN_PROCESSING_TIME,
    #                     EXPORT_COLUMN_DIFFERENCE_VCNC,
    #                     EXPORT_COLUMN_PLAATNUMMER,
    #                     EXPORT_COLUMN_LASER
    i = range(row[IMPORT_COLUMN_AANTALPLATEN])
    # print(row)
    # print(row[IMPORT_COLUMN_GROSS_TIME])
    for n in i:
        new_row = {EXPORT_COLUMN_PROGRAMMANUMMER: row[IMPORT_COLUMN_PROGRAMMANUMMER],
                    EXPORT_COLUMN_AANTALPLATEN: n,
                    EXPORT_COLUMN_MATERIAALCODE: row[IMPORT_COLUMN_MATERIAALCODE],
                    EXPORT_COLUMN_STUKS: row[IMPORT_COLUMN_STUKS],
                    EXPORT_COLUMN_TIJDPERPLAAT: row[IMPORT_COLUMN_TIJDPERPLAAT],
                    EXPORT_COLUMN_TOTALEINSTEKEN: row[IMPORT_COLUMN_TOTALEINSTEKEN],
                    EXPORT_COLUMN_TOTALESNIJLENGTE: row[IMPORT_COLUMN_TOTALESNIJLENGTE],
                    EXPORT_COLUMN_TIJD_METER: row[IMPORT_COLUMN_TIJD_METER],
                    EXPORT_COLUMN_GROSS_TIME: row[IMPORT_COLUMN_GROSS_TIME],
                    EXPORT_COLUMN_NETTO_TIME: row[IMPORT_COLUMN_NETTO_TIME],
                    EXPORT_COLUMN_PROCESSING_TIME: row[IMPORT_COLUMN_PROCESSING_TIME],
                    EXPORT_COLUMN_DIFFERENCE_VCNC: row[IMPORT_COLUMN_DIFFERENCE_VCNC],
                    EXPORT_COLUMN_PLAATNUMMER: row[IMPORT_COLUMN_PLAATNUMMER],
                    EXPORT_COLUMN_LASER: row[IMPORT_COLUMN_LASER]}
        
        data.loc[len(data)+1] = new_row

# # add total time and material found in ERP sheet to all the programnumbers found in ERP sheet
# for index, row in excel_sheet1.iterrows():
#     # index for knowing where to find the data with the correct programnumber in the ERP sheet
#     placementIndex1 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]].tolist()
#     # add time to existing time in the data set
#     data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] + row[IMPORT_COLUMN_ERP_TIJD]
#     # add time to existing time in the data set
#     data.loc[placementIndex1, EXPORT_COLUMN_ERP_STUKS] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_STUKS] + row[IMPORT_COLUMN_ERP_STUKS]
#     # add material code to the program number
#     data.loc[placementIndex1, EXPORT_COLUMN_ERP_MATERIAAL] = row[IMPORT_COLUMN_ERP_MATERIAAL]

# variables used for finding all usable data is Laser sheet
# lastProgramNumber = 0
# programmaNummer = 0
# LastGrossRunTime = 0
# GrossRunTime = 0
# LastPlaatNr = 0
# plaatNr = 0

# # fill empty cells with '' because python can't handle NaN
# # excel_sheet3 = excel_sheet3.fillna('')

# # clean up of the data, if there are rows that don't contain a gross time, plate and program number then delete these because the clog the algorithm
# # for index, row in excel_sheet3.iterrows():
# #     if row[IMPORT_COLUMN_LASER_TIJD] == '' and row[IMPORT_COLUMN_LASER_PLAAT] == '' and row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
# #         excel_sheet3.drop(index, axis='index', inplace=True)

# # variable used for debugging
# counted = 0

# run through all laser data and match programnumber to time and to the plate in the laser sheet,
# also match the data from the ERP to it if there is an programnumber in the ERP data that matches a programnumber in the laser data
# for index, row in excel_sheet1.iterrows():
    

# write dataframe with results to a new excel file
data.to_excel(FILE_NAME)

# read the excel file with the results
if DEBUGGING:
    result = pd.read_excel(FILE_NAME)
    print(result)
