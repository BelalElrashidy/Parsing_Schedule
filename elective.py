import json
import pandas as pd


json_file = 'elective.json'
excel_file = 'elective.xlsx'
xls = pd.ExcelFile(excel_file)
sheet_names = xls.sheet_names
Data_set = []
for sheet in sheet_names:
    data_frame = pd.read_excel(excel_file,sheet_name=sheet)
    column_name = data_frame.columns
    number_rows = len(data_frame)
    for i in range (1,9):
        j = 1
        print(f"{data_frame[column_name[i]][0]}")
        while j<number_rows:
            time = data_frame[column_name[0]][j]
            cell_value = data_frame[column_name[i]][j]
            if pd.isnull(cell_value):
                j+=1
                continue
            if pd.isnull(time):
                j +=2
                time = data_frame[column_name[0]][j]
            time = str(time)
            cell_value = str(cell_value)
            data = time+" _ " + cell_value
            j+=1
            Data_set.append(data)
    with open(json_file,'w') as file:
        json.dump(Data_set,file ,indent= 4)
        file.write('\n')
