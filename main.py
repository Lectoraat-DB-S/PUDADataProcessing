import pandas as pd
import numpy as np
from openpyxl import load_workbook

def Insert_row_(row_number, df, row_value):
    # Slice the upper half of the dataframe
    df1 = df[0:row_number]
  
    # Store the result of lower half of the dataframe
    df2 = df[row_number:]
  
    # Insert the row in the upper half dataframe
    df1.loc[row_number]=row_value
  
    # Concat the two dataframes
    df_result = pd.concat([df1, df2])
  
    # Reassign the index labels
    df_result.index = [*range(df_result.shape[0])]
  
    # Return the updated dataframe
    return df_result

file_location = "Data.xlsx"
file_name = "result.xlsx"

excel_sheet1 = pd.read_excel(file_location, sheet_name="ERP")
excel_sheet2 = pd.read_excel(file_location, sheet_name="Laser4")
# print(excel_sheet1)

poh = [20, 21, 22, 23, 33]

data = pd.DataFrame(columns=['ERP Programma Nummer', 'ERP TijdBonPerPlaat', 'Laser 4 GrossRunTime', 'Laser 4 Plaatnr', 'Laser 4 ProgrammaNaam'])
# data.loc[0] = 'lisa', 1, 2, 3, 4

index_new_data = 0
last_programmanummer = 0

for index, row in excel_sheet1.iterrows():
    # print(row['Programmanummer'])
    if row['Programmanummer'] != last_programmanummer:
        last_programmanummer = row['Programmanummer']
        data.loc[index_new_data] = row['Programmanummer'], 0, 0, 0, 0
        index_new_data = index_new_data + 1

# print(data)

for index, row in excel_sheet1.iterrows():
    # data.loc[data['Programma Nummer'] == row['Programmanummer'], data['ERP TijdBonPerPlaat']] = data['ERP TijdBonPerPlaat'] + row['TijdBonPerPlaat']
    placementIndex1 = data.index[data['ERP Programma Nummer'] == row['Programmanummer']].tolist()
    data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] = data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] + row['TijdBonPerPlaat']

lastProgramNumber = 0
programmaNummer = 0
LastGrossRunTime = 0
GrossRunTime = 0
FirstGrossRunTime = 0
LastPlaatNr = 0
plaatNr = 0

# print(excel_sheet2)
excel_sheet2 = excel_sheet2.fillna('')
# print(excel_sheet2)

index_new_data2 = 230

print(data)

counted = 0
count_the_amount = 9999 + 2

for index, row in excel_sheet2.iterrows():
    print("start")
    if lastProgramNumber == 0:
        lastProgramNumber = row['ProgrammaNaam']
    # if lastFoundProgramNumber == row['ProgrammaNaam']:
    #     placementIndex2 = data.index[data['ERP Programma Nummer'] == row['ProgrammaNaam']].tolist()

    # programmaNummer = row['ProgrammaNaam']
    # GrossRunTime = row['GrossRunTime']
    # plaatNr = row['PlaatNr']
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
            # placementIndex2 = data.index[data['ERP Programma Nummer'] == row['ProgrammaNaam']].tolist()
            placementIndex2 = data.index[data['ERP Programma Nummer'] == lastProgramNumber].tolist()
            # data = Insert_row_(placementIndex2, data, (data.loc[placementIndex2, 'ERP Programma Nummer'], data.loc[placementIndex2, 'ERP TijdBonPerPlaat'], GrossRunTime, plaatNr, programmaNummer))
            # data.loc[placementIndex2, 'ERP TijdBonPerPlaat'] = data.loc[placementIndex2, 'ERP Programma Nummer'] + row['TijdBonPerPlaat']
            # data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] = data.loc[placementIndex1, 'ERP TijdBonPerPlaat'] + row['TijdBonPerPlaat']
            # data.loc[index_new_data2] = data.loc[placementIndex2, 'ERP Programma Nummer'], data.loc[placementIndex2, 'ERP TijdBonPerPlaat'], GrossRunTime, plaatNr, programmaNummer
            print(placementIndex2)
            if not placementIndex2:
                # if 
                print('list is empty')
                test1 = ''
                test2 = ''
                # print(placementIndex2)
                poh = [0,0,0,0,0]
            else:
                test1 = data.iat[placementIndex2[0], 0]
                test2 = data.iat[placementIndex2[0], 1]
            if 1:
                # print(placementIndex2[0])
                # test1 = data.loc[placementIndex2, 'ERP Programma Nummer']
                # -----test1 = data.iat[placementIndex2[0], 0]
                # test1 = data.loc[placementIndex2, 0]
                # test2 = data.loc[placementIndex2, 'ERP TijdBonPerPlaat']
                # -----test2 = data.iat[placementIndex2[0], 1]
                # test2 = data.loc[placementIndex2, 1]
                test3 = GrossRunTime
                test4 = plaatNr
                test5 = programmaNummer
                print("fuck")
                print(row)
                print("dafuck")
                print(test1)
                print("shit")
                print(test2)
                print("baka")
                print(test3)
                print(test4)
                print(test5)
                print(str(test1) + "__" + str(test2) + "__"  + str(test3) + "__"  + str(test4) + "__"  + str(test5))
                new_row = {'ERP Programma Nummer': test1, 'ERP TijdBonPerPlaat': test2, 'Laser 4 GrossRunTime': test3, 'Laser 4 Plaatnr': test4, 'Laser 4 ProgrammaNaam': test5}
                data.loc[len(data)+1] = new_row
                # data = data.append(new_row, ignore_index=True)

                # data = data.reset_index(drop=True)
            # index_new_data2 = index_new_data2 + 1
    else:
        print("not a matched number")
        # test1 = data.iat[placementIndex2[0], 0]
        # test2 = data.iat[placementIndex2[0], 1]
        # test3 = GrossRunTime
        # test4 = plaatNr
        # test5 = programmaNummer
        # new_row = {'ERP Programma Nummer': test1, 'ERP TijdBonPerPlaat': test2, 'Laser 4 GrossRunTime': test3, 'Laser 4 Plaatnr': test4, 'Laser 4 ProgrammaNaam': test5}
        # data.loc[len(data)+1] = new_row
        lastProgramNumber = row['ProgrammaNaam']
        LastPlaatNr = 0

    print(lastProgramNumber)
    print(row['ProgrammaNaam'])
    print(LastPlaatNr)
    print(row['PlaatNr'])
    print(row['GrossRunTime'])

    if counted == count_the_amount:
        break
    else:
        counted = counted + 1




data.to_excel(file_name)

result = pd.read_excel(file_name)
print(result)
