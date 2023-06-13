import pandas as pd

excel_file = 'Core.xlsx'

# Create an ExcelFile object
xls = pd.ExcelFile(excel_file)

# Get the sheet names
sheet_names = xls.sheet_names

# Print the sheet names
for sheet_name in sheet_names:
    print(sheet_name)
