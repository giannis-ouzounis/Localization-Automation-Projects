# -*- coding: utf-8 -*-
"""
Loads a specified .json file into a pandas DataFrame.
Adds an ID column (row index) and shifts it to column A.
Saves the reordered DataFrame as an Excel file (*_output.xlsx) in the same folder.
Reads the new Excel file back in and prints it—confirming layout ready for translation workflow.
"""
import os
import pandas as pd
import json

# Load JSON data
file_path_json = r'C:\Users\stephanie.pietz\Documents\Projects\ORMCO_2401_P0006\localization 1.json'

with open(file_path_json, encoding='utf-8') as f:
    data = json.load(f)

# Create a DataFrame using pandas
df = pd.DataFrame(data)

# Add a new column with the ID for each language
df['ID'] = df.index

# Reorder columns
df = df[['ID'] + [col for col in df.columns if col != 'ID']]

# Get the directory of the JSON file
output_directory = os.path.dirname(file_path_json)

# Construct the output Excel file path
file_name_without_extension = os.path.splitext(os.path.basename(file_path_json))[0]
file_path_excel = os.path.join(output_directory, f'{file_name_without_extension}_output.xlsx')

# Write to Excel file
df.to_excel(file_path_excel, index=False)

# Open the created Excel file
df_read = pd.read_excel(file_path_excel)
print(df_read)
