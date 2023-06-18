import pandas as pd
import json

json_file = 'Core.json'
excel_file = 'Core.xlsx'
Sheet_name = 'SUMMER 2023(BS1BS2MS1)'
data_frame = pd.read_excel(excel_file,sheet_name=Sheet_name)
column_name = data_frame.columns
Data_set = []
number_rows = len(data_frame)
for j in range(1,len(column_name),1):

    for i in range(number_rows):
        time = data_frame[column_name[0]][i]
        cell_value = data_frame[column_name[j]][i]
        if pd.isnull(cell_value):
            continue
        if pd.isnull(time) and i>1:
            test = data_frame[column_name[0]][i-1]
            if pd.isnull(test):
                time = data_frame[column_name[0]][i-2]
            else:
                time = test
        
        if not pd.isnull(time):
            time = str(time)
            cell_value = str(cell_value)
            data = time + ' _ '+ cell_value
        else:
            cell_value = str(cell_value)
            data =  ' _ '+ cell_value
        Data_set.append(data)
print()
with open(json_file,'w') as file:
    json.dump(Data_set,file, indent=4)