import pandas as pd
import numpy as np
from openpyxl import load_workbook

file_location = "Data.xlsx"
file_name = "result2.xlsx"

excel_sheet1 = pd.read_excel(file_location, sheet_name="ERP")
# excel_sheet2 = pd.read_excel(file_location, sheet_name="Laser 3")
excel_sheet2 = pd.read_excel(file_location, sheet_name="Laser 4")

poh = [20, 21, 22, 23, 33]

# data = pd.DataFrame(columns=['ERP Programma Nummer', 'ERP TijdBonPerPlaat', 'Laser 3 GrossRunTime', 'Laser 3 Plaatnr', 'Laser 3 ProgrammaNaam'])
data = pd.DataFrame(columns=['ERP Programma Nummer', 'ERP TijdBonPerPlaat', 'Laser 4 GrossRunTime', 'Laser 4 Plaatnr', 'Laser 4 ProgrammaNaam'])

index_new_data = 0
last_programmanummer = 0

for index, row in excel_sheet1.iterrows():
    if row['Programmanummer'] != last_programmanummer:
        last_programmanummer = row['Programmanummer']
        data.loc[index_new_data] = row['Programmanummer'], 0, 0, 0, 0
        index_new_data = index_new_data + 1

for index, row in excel_sheet1.iterrows():
    placementIndex1 = data.index[data['ERP Programma Nummer'] == row['Programmanummer']].tolist()
    data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] = data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] + row['TijdBonPerPlaat']

lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
FirstGrossRunTime = 0
LastPlaatNr = 0
plaatNr = 0

excel_sheet2 = excel_sheet2.fillna('')

index_new_data2 = 230

counted = 0
count_the_amount = 9999 + 2

for index, row in excel_sheet2.iterrows():
    print("start")
    if lastProgramNumber == 0:
        lastProgramNumber = row['ProgrammaNaam']
    print("check for number")
    if lastProgramNumber == row['ProgrammaNaam'] or row['ProgrammaNaam'] == '':
        print("matched number")
        print("check for plaatnr")
        if row['PlaatNr'] == LastPlaatNr or row['PlaatNr'] == '':
            print("matched plaatnr")
            print("check for grossruntime")
            if row['GrossRunTime'] != '':
                print("matched grossruntime")
                LastGrossRunTime = row['GrossRunTime']
        else:
            print("not a matched plaatnr")
            plaatNr = LastPlaatNr
            LastPlaatNr = row['PlaatNr']
            programmaNummer = lastProgramNumber
            GrossRunTime = LastGrossRunTime
            LastGrossRunTime = row['GrossRunTime']
            placementIndex2 = data.index[data['ERP Programma Nummer'] == lastProgramNumber].tolist()
            if not placementIndex2:
                # if 
                print('list is empty')
                test1 = ''
                test2 = ''
                poh = [0,0,0,0,0]
            else:
                test1 = data.iat[placementIndex2[0], 0]
                test2 = data.iat[placementIndex2[0], 1]
            if 1:
                test3 = GrossRunTime
                test4 = plaatNr
                test5 = programmaNummer
                # new_row = {'ERP Programma Nummer': test1, 'ERP TijdBonPerPlaat': test2, 'Laser 3 GrossRunTime': test3, 'Laser 3 Plaatnr': test4, 'Laser 3 ProgrammaNaam': test5}
                new_row = {'ERP Programma Nummer': test1, 'ERP TijdBonPerPlaat': test2, 'Laser 4 GrossRunTime': test3, 'Laser 4 Plaatnr': test4, 'Laser 4 ProgrammaNaam': test5}
                data.loc[len(data)+1] = new_row
    else:
        print("not a matched number")
        lastProgramNumber = row['ProgrammaNaam']
        LastPlaatNr = 0

    if counted == count_the_amount:
        break
    else:
        counted = counted + 1




data.to_excel(file_name)

result = pd.read_excel(file_name)
print(result)
