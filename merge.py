import pandas as pd
import os

# Specify the folder path containing XLSX files
folder_path = 'C:/Users/Eduarda Branco/Documents/All'

# Create an empty list to store DataFrames
dataframes = []

# List all files in the folder
all_files = os.listdir(folder_path)

# Filter files to include only XLSX files
xlsx_files = [file for file in all_files if file.endswith('.xlsx')]

for excel_file in xlsx_files:
    excel_file_path = os.path.join(folder_path, excel_file)
    
    # Get the list of sheet names in the Excel file
    with pd.ExcelFile(excel_file_path) as xls:
        sheet_names = xls.sheet_names
    
    for sheet_name in sheet_names:
        # Read each sheet and store the DataFrame in the list
        df = pd.read_excel(excel_file_path, sheet_name)
        dataframes.append(df)

# Concatenate the list of DataFrames into a single DataFrame
merged_data = pd.concat(dataframes, ignore_index=True)

# Save the merged data to a new Excel file
merged_data.to_excel('merged_file.xlsx', index=False)

print("Merging and saving completed.")