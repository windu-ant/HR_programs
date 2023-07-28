import pandas as pd
import os
import shutil
from datetime import datetime


# Set the CSV and excel file
df_csv = pd.read_csv('export.csv')
df_excel = pd.read_excel('HireRight_List.xlsx')

# Filter rows in CSV that start with "Completed"
df_csv_completed = df_csv[df_csv['Tab'].str.startswith('Completed')].copy()

# Find the rows in the CSV that are not present in the Excel file based on the "Request #" column
df_difference = df_csv_completed[~df_csv_completed['Request #'].isin(df_excel['Request #'])]

# Append the different rows to the Excel DataFrame
df_new_excel = pd.concat([df_excel, df_difference])

# Write the new DataFrame to an Excel file
df_new_excel.to_excel('new_data.xlsx', index=False)

# Get the time
now = datetime.now()
# Set time string
now_str = now.strftime("%Y_%m_%d_%H%M%S")

# Get CWD
cwd = os.getcwd()
source = os.path.join(cwd, 'HireRight_List.xlsx')

# Set the new file name with date time
destination = os.path.join(cwd, '_Backup', f'HireRight_List_{now_str}.xlsx')

# Move the file
dest = shutil.move(source, destination)

# Finally rename the new combined data to list
os.rename('new_data.xlsx', 'HireRight_List.xlsx')

print("Done")