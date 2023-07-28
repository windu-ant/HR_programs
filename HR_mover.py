import pandas as pd
import os
import shutil

# Load the excel data
df = pd.read_excel('HireRight_List.xlsx')

# Loop over the DataFrame
for index, row in df.iterrows():
    # Check the "user group" column for each row
    if pd.isna(row['User Group']) or row['User Group'] != 'x':
        if pd.isna(row['Middle Name']):
            # If 'Middle Name' is NaN, only use 'First Name' and 'Last Name'
            name = f"{row['Last Name']}, {row['First Name']}"
            df.at[index, 'User Group'] = 'x'
        else:
            # If 'Middle Name' is not NaN, use 'First Name', 'Middle Name', and 'Last Name'
            name = f"{row['Last Name']}, {row['First Name']} {row['Middle Name']}"
            df.at[index, 'User Group'] = 'x'


        # set current working directoy to move files to
        cwd = os.getcwd()
        # one directory up
        parent_directory = os.path.join(cwd, '..')
        # sent backgrounds folder
        sent_backgrounds = os.path.join(parent_directory, '1 - Sent for Background check')

        # Search for a folder named as such
        found = False
        for root, dirs, files in os.walk(sent_backgrounds):
            matching_dirs = [d for d in dirs if d.startswith(name)]
            for dir in matching_dirs:
                found = True
                # Move the folder into the current working directory
                moved_file = shutil.move(os.path.join(root, dir), os.getcwd())
                print(f"Moved {moved_file}\n")

        if not found:
            file_not_found = os.path.join(os.getcwd(), f"{name}_NOT_FOUND")
            os.mkdir(file_not_found)
            print(f"NOT FOUND {file_not_found}\n")
            

# Save the updated DataFrame to the Excel file
df.to_excel('HireRight_List.xlsx', index=False)
