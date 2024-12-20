import pandas as pd
import numpy as np
from openpyxl import load_workbook

# setup constants
# put a 1 if debugging should be enabled
DEBUGGING = 0
MACHINE_NUMBER = 3
NEW_DATA = 1
if MACHINE_NUMBER != 3 and MACHINE_NUMBER != 4:
    print("wrong machine number used")
    print("acceptable machine numbers are 3 and 4")
    print("shutting down, toedeloe")
    exit()

# declare constants
if NEW_DATA:
    FILE_LOCATION = "Data_export_alles_vanaf_2023.xlsx"
else:
    FILE_LOCATION = "Data.xlsx"
if MACHINE_NUMBER == 3:
    if NEW_DATA:
        FILE_NAME = "result_laser_3_vanaf_2023.xlsx"
    else:
        FILE_NAME = "resultLaser3V5.xlsx"
elif MACHINE_NUMBER == 4:
    if NEW_DATA:
        FILE_NAME = "result_laser_4_vanaf_2023.xlsx"
    else:
        FILE_NAME = "resultLaser4V5.xlsx"

# sheet names of the imported file
if NEW_DATA:
    EXCEL_SHEET_NAME1 = "RidderiQ"
else:
    EXCEL_SHEET_NAME1 = "ERP"
if NEW_DATA:
    EXCEL_SHEET_NAME2 = "WiCAM"
else:
    EXCEL_SHEET_NAME2 = "WICAM"
if MACHINE_NUMBER == 3:
    if NEW_DATA:
        EXCEL_SHEET_NAME3 = "ThingsBoard laser 3"
    else:
        EXCEL_SHEET_NAME3 = "Laser 3"
elif MACHINE_NUMBER == 4:
    if NEW_DATA:
        EXCEL_SHEET_NAME3 = "ThingsBoard laser 4"
    else:
        EXCEL_SHEET_NAME3 = "Laser4"

# names of the columns from the imported file
# sheet 1 ERP
IMPORT_COLUMN_ERP_PROGRAMMANUMMER = 'Programmanummer'
IMPORT_COLUMN_ERP_MATERIAAL = 'MateriaalCode'
IMPORT_COLUMN_ERP_ORDER_BON = 'OrderBon'
IMPORT_COLUMN_ERP_STUKS = 'StuksPerPlaat'
if NEW_DATA:
    IMPORT_COLUMN_ERP_TIJD = 'TijdBonPerPlaat (s)'
else:
    IMPORT_COLUMN_ERP_TIJD = 'TijdBonPerPlaat'
# sheet 2 WICAM
IMPORT_COLUMN_WICAM_PROGRAMMANUMMER = 'Programmanummer'
IMPORT_COLUMN_WICAM_ORDER_BON = 'OrderBon'
if NEW_DATA:
    IMPORT_COLUMN_WICAM_SNIJLENGTE = 'SnijlengtePerStuk (mm)'
else:
    IMPORT_COLUMN_WICAM_SNIJLENGTE = 'SnijlengtePerStuk'
# sheet 3 LASER
IMPORT_COLUMN_LASER_STARTTIME = 'Timestamp'
IMPORT_COLUMN_LASER_TIJD = 'GrossRunTime'
IMPORT_COLUMN_LASER_PLAAT = 'PlaatNr'
IMPORT_COLUMN_LASER_PROGRAMMANUMMER = 'ProgrammaNaam'
IMPORT_COLUMN_LASER_NETPROCESSINGTIME = 'NetProcessingTime'
IMPORT_COLUMN_LASER_NETRUNTIME = 'NetRunTime'

# names of the exported columns
EXPORT_COLUMN_ERP_PROGRAMMANUMMER = 'ERP Programma Nummer'
if NEW_DATA:
    EXPORT_COLUMN_ERP_TIJD = 'ERP TijdBonPerPlaat (s)'
else:
    EXPORT_COLUMN_ERP_TIJD = 'ERP TijdBonPerPlaat'
EXPORT_COLUMN_ERP_MATERIAAL = 'ERP MateriaalCode'
EXPORT_COLUMN_ERP_STUKS = 'ERP Stuks'
if NEW_DATA:
    EXPORT_COLUMN_WICAM_AVG_TIMEDISTANCE = 'Average time per distance (s/m)'
else:
    EXPORT_COLUMN_WICAM_AVG_TIMEDISTANCE = 'Average time per distance'
if MACHINE_NUMBER == 3:
    EXPORT_COLUMN_LASER_TIJD = 'Laser 3 GrossRunTime'
    EXPORT_COLUMN_LASER_PLAAT = 'Laser 3 Plaatnr'
    EXPORT_COLUMN_LASER_PROGRAMMANUMMER = 'Laser 3 ProgrammaNaam'
    EXPORT_COLUMN_LASER_DIFFERENCE = 'Laser 3 Actual Time Difference'
    EXPORT_COLUMN_LASER_TIJD_NET_PROCESSING = 'Laser 3 NetProcessingTime'
    EXPORT_COLUMN_LASER_TIJD_NET_RUN = 'Laser 3 NetRunTime'
    EXPORT_COLUMN_LASER_STARTTIME = 'Laser 3 Start time'
elif MACHINE_NUMBER == 4:
    EXPORT_COLUMN_LASER_TIJD = 'Laser 4 GrossRunTime'
    EXPORT_COLUMN_LASER_PLAAT = 'Laser 4 Plaatnr'
    EXPORT_COLUMN_LASER_PROGRAMMANUMMER = 'Laser 4 ProgrammaNaam'
    EXPORT_COLUMN_LASER_DIFFERENCE = 'Laser 4 Actual Time Difference'
    EXPORT_COLUMN_LASER_TIJD_NET_PROCESSING = 'Laser 4 NetProcessingTime'
    EXPORT_COLUMN_LASER_TIJD_NET_RUN = 'Laser 4 NetRunTime'
    EXPORT_COLUMN_LASER_STARTTIME = 'Laser 4 Start time'

# testing constant
COUNT_THE_AMOUNT = 120 + 2

# ERP excel sheet
excel_sheet1 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME1, converters={IMPORT_COLUMN_ERP_ORDER_BON: str})
# WICAM excel sheet
excel_sheet2 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME2)
# Laser excel sheet
excel_sheet3 = pd.read_excel(FILE_LOCATION, sheet_name=EXCEL_SHEET_NAME3)
# dataframe that will be filled we the results
data = pd.DataFrame(columns=[EXPORT_COLUMN_ERP_PROGRAMMANUMMER, EXPORT_COLUMN_ERP_TIJD, EXPORT_COLUMN_ERP_MATERIAAL, EXPORT_COLUMN_ERP_STUKS, EXPORT_COLUMN_WICAM_AVG_TIMEDISTANCE, EXPORT_COLUMN_LASER_DIFFERENCE, EXPORT_COLUMN_LASER_TIJD, EXPORT_COLUMN_LASER_TIJD_NET_PROCESSING, EXPORT_COLUMN_LASER_TIJD_NET_RUN, EXPORT_COLUMN_LASER_PLAAT, EXPORT_COLUMN_LASER_PROGRAMMANUMMER, EXPORT_COLUMN_LASER_STARTTIME])

excel_sheet1[IMPORT_COLUMN_ERP_ORDER_BON] = excel_sheet1[IMPORT_COLUMN_ERP_ORDER_BON].str.replace('.', '-')

# indexing and checking for removing all duplicate programnumbers
index_new_data = 0
last_programmanummer = 0

# Find all programnumbers that exist in the ERP sheet
for index, row in excel_sheet1.iterrows():
    if row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] != last_programmanummer:
        last_programmanummer = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]
        data.loc[index_new_data] = row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER], 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        index_new_data = index_new_data + 1

# add total time and material found in ERP sheet to all the programnumbers found in ERP sheet
for index, row in excel_sheet1.iterrows():
    # index for knowing where to find the data with the correct programnumber in the ERP sheet
    placementIndex1 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_ERP_PROGRAMMANUMMER]].tolist()
    # add time to existing time in the data set
    data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_TIJD] + row[IMPORT_COLUMN_ERP_TIJD]
    # add time to existing time in the data set
    data.loc[placementIndex1, EXPORT_COLUMN_ERP_STUKS] = data.loc[placementIndex1, EXPORT_COLUMN_ERP_STUKS] + row[IMPORT_COLUMN_ERP_STUKS]
    # add material code to the program number
    data.loc[placementIndex1, EXPORT_COLUMN_ERP_MATERIAAL] = row[IMPORT_COLUMN_ERP_MATERIAAL]

# variables used for finding all usable data is Laser sheet
lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
LastNetProcessingTime = 0
NetProcessingTime = 0
LastNetRunTime = 0
NetRunTime = 0
LastPlaatNr = 0
plaatNr = 0
startTime = 0
lastStartTime = 0
changeStartTime = 1

# fill empty cells with '' because python can't handle NaN
excel_sheet3 = excel_sheet3.fillna('')

# clean up of the data, if there are rows that don't contain a gross time, plate and program number then delete these because the clog the algorithm
for index, row in excel_sheet3.iterrows():
    if row[IMPORT_COLUMN_LASER_TIJD] == '' and row[IMPORT_COLUMN_LASER_NETPROCESSINGTIME] == '' and row[IMPORT_COLUMN_LASER_NETRUNTIME] == '' and row[IMPORT_COLUMN_LASER_PLAAT] == '' and row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER] == '':
        excel_sheet3.drop(index, axis='index', inplace=True)

# variable used for debugging
counted = 0

# run through all laser data and match programnumber to time and to the plate in the laser sheet,
# also match the data from the ERP to it if there is an programnumber in the ERP data that matches a programnumber in the laser data
for index, row in excel_sheet3.iterrows():
    if DEBUGGING:
        print("start")
    # record start time of listing
    if changeStartTime == 1:
        lastStartTime = row[IMPORT_COLUMN_LASER_STARTTIME]
        changeStartTime = 0
    # for the first run safe a number to prevent skipping
    if lastProgramNumber == 0:
        lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
        # lastStartTime = row[IMPORT_COLUMN_LASER_STARTTIME]
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
            if row[IMPORT_COLUMN_LASER_NETPROCESSINGTIME] != '':
                if DEBUGGING:
                    print("matched netprocessingtime")
                LastNetProcessingTime = row[IMPORT_COLUMN_LASER_NETPROCESSINGTIME]
            if row[IMPORT_COLUMN_LASER_NETRUNTIME] != '':
                if DEBUGGING:
                    print("matched netruntime")
                LastNetRunTime = row[IMPORT_COLUMN_LASER_NETRUNTIME]
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
            NetProcessingTime = LastNetProcessingTime
            LastNetProcessingTime = row[IMPORT_COLUMN_LASER_NETPROCESSINGTIME]
            NetRunTime = LastNetRunTime
            LastNetRunTime = row[IMPORT_COLUMN_LASER_NETRUNTIME]
            startTime = lastStartTime
            changeStartTime = 1

            # index for knowing where to find the data with the same programnumber in the ERP sheet
            placementIndex2 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == lastProgramNumber].tolist()

            # in case the same programnumber from the laser sheet could not be found in the ERP sheet, leave cells empty
            if not placementIndex2:
                if DEBUGGING:
                    print('list is empty')
                ERPProgrammanummer = ''
                ERPTijd = ''
                ERPMateriaal = ''
                ERPStuks = ''
            # in case the same programnumber was found, fill the cells with the correct data
            else:
                ERPProgrammanummer = data.iat[placementIndex2[0], 0]
                ERPTijd = data.iat[placementIndex2[0], 1]
                ERPMateriaal = data.iat[placementIndex2[0], 2]
                ERPStuks = data.iat[placementIndex2[0],3]

            # calculate the time difference between expected and actual
            if ERPTijd != '' and GrossRunTime != '':
                # print(programmaNummer)
                # print(ERPProgrammanummer)
                # print(type(GrossRunTime))
                # print(GrossRunTime)
                # print(type(ERPTijd))
                # print(ERPTijd)
                timeDifference = GrossRunTime - ERPTijd
            elif ERPTijd == '':
                timeDifference = 'No ERP time'
            elif GrossRunTime == '':
                timeDifference = 'No GrossRunTime'

            # variable to safe the average in
            avgTimeDistance = 0
            totalSnijlengte = 0
            wicamCheck = 0
            # add average time per distance
            # check if the programnumber is present in the ERP, otherwise we don't know how many pieces there are in a plate
            if ERPProgrammanummer != '':
                # run through the WICAM sheet
                for index, row in excel_sheet2.iterrows():
                    #  check to see if there is a matching programnumber with the laser machine
                    if row[IMPORT_COLUMN_WICAM_PROGRAMMANUMMER] == ERPProgrammanummer:
                        placementIndex3 = excel_sheet1.index[excel_sheet1[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_WICAM_PROGRAMMANUMMER]].tolist()
                        for i in placementIndex3:
                            if row[IMPORT_COLUMN_WICAM_ORDER_BON] == excel_sheet1.loc[i, IMPORT_COLUMN_ERP_ORDER_BON]:
                                wicamCheck = 1
                                totalSnijlengte = totalSnijlengte + ((row[IMPORT_COLUMN_WICAM_SNIJLENGTE]/1000)*excel_sheet1.loc[i, IMPORT_COLUMN_ERP_STUKS])
                if wicamCheck:
                    avgTimeDistance = ERPTijd/totalSnijlengte
                else:
                    avgTimeDistance = 'No Wicam programnumber'
            else:
                avgTimeDistance = 'No ERP programnumber'

            # create the new row that will have to be added to the dataframe
            new_row = {EXPORT_COLUMN_ERP_PROGRAMMANUMMER: ERPProgrammanummer,
                       EXPORT_COLUMN_ERP_TIJD: ERPTijd,
                       EXPORT_COLUMN_ERP_MATERIAAL: ERPMateriaal,
                       EXPORT_COLUMN_ERP_STUKS: ERPStuks,
                       EXPORT_COLUMN_LASER_DIFFERENCE: timeDifference,
                       EXPORT_COLUMN_WICAM_AVG_TIMEDISTANCE: avgTimeDistance,
                       EXPORT_COLUMN_LASER_TIJD: GrossRunTime,
                       EXPORT_COLUMN_LASER_TIJD_NET_PROCESSING: NetProcessingTime,
                       EXPORT_COLUMN_LASER_TIJD_NET_RUN: NetRunTime,
                       EXPORT_COLUMN_LASER_PLAAT: plaatNr,
                       EXPORT_COLUMN_LASER_PROGRAMMANUMMER: programmaNummer,
                       EXPORT_COLUMN_LASER_STARTTIME: startTime}
            
            # add the new row to the dataframe
            data.loc[len(data)+1] = new_row
    # the same program number wasn't found
    else:
        if DEBUGGING:
            print("not a matched program number")
        # if it doesn't have a time then we don't need to safe it
        if LastGrossRunTime == '' and LastNetProcessingTime == '' and LastNetRunTime == '':
            if DEBUGGING:
                print("No time was saved for this platenr and programnr")
            lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
            if row[IMPORT_COLUMN_LASER_PLAAT] != '':
                LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            else:
                LastPlaatNr = 0
            LastGrossRunTime = 0
        # if it does have a time then we need to safe it
        elif LastGrossRunTime != '' or LastNetProcessingTime != '' or LastNetRunTime != '':
            if DEBUGGING:
                print("It does contain time")
            plaatNr = LastPlaatNr
            # in case that we continue with a plate that isn't 0
            if row[IMPORT_COLUMN_LASER_PLAAT] != '':
                LastPlaatNr = row[IMPORT_COLUMN_LASER_PLAAT]
            else:
                LastPlaatNr = 0
            programmaNummer = lastProgramNumber
            lastProgramNumber = row[IMPORT_COLUMN_LASER_PROGRAMMANUMMER]
            GrossRunTime = LastGrossRunTime
            LastGrossRunTime = row[IMPORT_COLUMN_LASER_TIJD]
            NetProcessingTime = LastNetProcessingTime
            LastNetProcessingTime = row[IMPORT_COLUMN_LASER_NETPROCESSINGTIME]
            NetRunTime = LastNetRunTime
            LastNetRunTime = row[IMPORT_COLUMN_LASER_NETRUNTIME]
            startTime = lastStartTime
            changeStartTime = 1

            # index for knowing where to find the data with the same programnumber in the ERP sheet
            placementIndex2 = data.index[data[EXPORT_COLUMN_ERP_PROGRAMMANUMMER] == programmaNummer].tolist()

            # in case the same programnumber from the laser sheet could not be found in the ERP sheet, leave cells empty
            if not placementIndex2:
                if DEBUGGING:
                    print('list is empty')
                ERPProgrammanummer = ''
                ERPTijd = ''
                ERPMateriaal = ''
                ERPStuks = ''
            # in case the same programnumber was found, fill the cells with the correct data
            else:
                ERPProgrammanummer = data.iat[placementIndex2[0], 0]
                ERPTijd = data.iat[placementIndex2[0], 1]
                ERPMateriaal = data.iat[placementIndex2[0], 2]
                ERPStuks = data.iat[placementIndex2[0],3]

            # calculate the time difference between expected and actual
            if ERPTijd != '' and GrossRunTime != '':
                timeDifference = GrossRunTime - ERPTijd
            else:
                timeDifference = 'No ERP time'

            # variable to safe the average in
            avgTimeDistance = 0
            totalSnijlengte = 0
            wicamCheck = 0
            # add average time per distance
            # check if the programnumber is present in the ERP, otherwise we don't know how many pieces there are in a plate
            if ERPProgrammanummer != '':
                # run through the WICAM sheet
                for index, row in excel_sheet2.iterrows():
                    #  check to see if there is a matching programnumber with the laser machine
                    if row[IMPORT_COLUMN_WICAM_PROGRAMMANUMMER] == ERPProgrammanummer:
                        placementIndex3 = excel_sheet1.index[excel_sheet1[IMPORT_COLUMN_ERP_PROGRAMMANUMMER] == row[IMPORT_COLUMN_WICAM_PROGRAMMANUMMER]].tolist()
                        for i in placementIndex3:
                            if row[IMPORT_COLUMN_WICAM_ORDER_BON] == excel_sheet1.loc[i, IMPORT_COLUMN_ERP_ORDER_BON]:
                                wicamCheck = 1
                                totalSnijlengte = totalSnijlengte + ((row[IMPORT_COLUMN_WICAM_SNIJLENGTE]/1000)*excel_sheet1.loc[i, IMPORT_COLUMN_ERP_STUKS])
                if wicamCheck:
                    avgTimeDistance = ERPTijd/totalSnijlengte
                else:
                    avgTimeDistance = 'No Wicam programnumber'
            else:
                avgTimeDistance = 'No ERP programnumber'

            # create the new row that will have to be added to the dataframe
            new_row = {EXPORT_COLUMN_ERP_PROGRAMMANUMMER: ERPProgrammanummer,
                       EXPORT_COLUMN_ERP_TIJD: ERPTijd,
                       EXPORT_COLUMN_ERP_MATERIAAL: ERPMateriaal,
                       EXPORT_COLUMN_ERP_STUKS: ERPStuks,
                       EXPORT_COLUMN_LASER_DIFFERENCE: timeDifference,
                       EXPORT_COLUMN_WICAM_AVG_TIMEDISTANCE: avgTimeDistance,
                       EXPORT_COLUMN_LASER_TIJD: GrossRunTime,
                       EXPORT_COLUMN_LASER_TIJD_NET_PROCESSING: NetProcessingTime,
                       EXPORT_COLUMN_LASER_TIJD_NET_RUN: NetRunTime,
                       EXPORT_COLUMN_LASER_PLAAT: plaatNr,
                       EXPORT_COLUMN_LASER_PROGRAMMANUMMER: programmaNummer,
                       EXPORT_COLUMN_LASER_STARTTIME: startTime}
            
            # add the new row to the dataframe
            data.loc[len(data)+1] = new_row
        else:
            if DEBUGGING:
                print('Something went horribly wrong')
                print(LastGrossRunTime)
                print(LastNetProcessingTime)
                print(LastNetRunTime)

    # debugging functions which enables only part of the excel sheet from being checked
    if DEBUGGING:
        if counted == COUNT_THE_AMOUNT:
            break
        else:
            counted = counted + 1

# this removes the unnecessary lines at the beginning of the file that are empty
for index, row in data.iterrows():
    if row[EXPORT_COLUMN_LASER_PLAAT] == 0 and row[EXPORT_COLUMN_LASER_PROGRAMMANUMMER] == 0 and row[EXPORT_COLUMN_LASER_TIJD] == 0:
        data.drop(index, axis='index', inplace=True)

# write dataframe with results to a new excel file
data.to_excel(FILE_NAME)

# read the excel file with the results
if DEBUGGING:
    result = pd.read_excel(FILE_NAME)
    print(result)
