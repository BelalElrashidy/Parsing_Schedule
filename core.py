import pandas as pd
import json
from unmerge import unmerge
json_file = 'Core.json'
old = 'Core.xlsx'
excel_file = unmerge(old)
Sheet_name = 'SUMMER 2023(BS1BS2MS1)'
data_frame = pd.read_excel(excel_file,sheet_name=Sheet_name)
column_name = data_frame.columns
Data_set = {}
years = {}
Groups={}
year = str(column_name[1])
day = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SUNDAY","Saturday"]
number_rows = len(data_frame)
for j in range(1,len(column_name),1):
    counter = 0
    group = ""
    Days = {}
    Slots ={}
    periods = ["first","second","third","forht","fifth","sixth"]
    number = 0
    i = 0
    the_day ="MONDAY"
    
    while i < number_rows:
        if i >= number_rows:
            break
        one_slot= {"name":"","lecturer":"","room":"","from":"","to":"","type":""}
        time = data_frame[column_name[0]][i]
        cell_value = data_frame[column_name[j]][i]
        if pd.isnull(cell_value) and not time in day:
            i+=1
            continue    
        if time =='F' or pd.isnull(time):
            group = str(data_frame[column_name[j]][i])
            i+=2
        elif not pd.isnull(time):
            if cell_value in day:
                years[year] = Groups
                year = str(column_name[j+1])
                break
            elif time in day:
                Days[the_day] = Slots
                the_day = time
                Slots = {}
                number = 0
                i+=1
            elif cell_value == "Elective courses on Physical Education":
                one_slot["name"] = cell_value
                duration = time.split("-",1)
                one_slot["from"] = duration[0]
                one_slot["to"] = duration[1]
                one_slot["type"] = "Sports"
                Slots[periods[number]] = one_slot
                number+=1
                i+=2
            else:
                cell_value = str(cell_value)
                cell = cell_value.split("(",1)
                name = cell[0]
                splitted = cell[1].split(')',1)
                if not len(splitted[1]) == 0:
                    name+=splitted[1]
                type = splitted[0]
                lecturer =data_frame[column_name[j]][i+1]
                room = data_frame[column_name[j]][i+2]
                duration = time.split("-",1)
                one_slot["from"] = duration[0]
                one_slot["to"] = duration[1]
                one_slot["lecturer"] = lecturer
                one_slot["name"] = name
                one_slot["room"] = room
                one_slot ["type"] = type
                Slots[periods[number]] = one_slot                
                number+=1
                i+=3
    Groups[group] = Days
Data_set = years     
                
                    
    
with open(json_file,'w') as file:
    json.dump(Data_set,file, indent=4)
