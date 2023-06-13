import pandas as pd

excel_file = 'Core.xlsx'
Sheet_name = 'SUMMER 2023(BS1BS2MS1)'
data_frame = pd.read_excel(excel_file,sheet_name=Sheet_name)
# print(data_frame)
i=0
j=0
for row_index, row in data_frame.items():
    for col_index, cell_value in enumerate(row):
        print(f"Cell ({row_index+1}, {col_index+1}): {cell_value}")
