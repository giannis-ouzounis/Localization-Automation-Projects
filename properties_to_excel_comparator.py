# -*- coding: utf-8 -*-
"""
Reads two .properties files, removes empty lines and nbsp, imports them into separate Excel sheets, and creates a third sheet comparing them based on keys.
"""

import pandas as pd
from openpyxl import load_workbook

#Function to read the data from the properties files, splitting Key from the data/stings
def read_properties(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    data = {}
    for line in lines:
        if '=' in line:
            key, value = map(str.strip, line.split('=', 1))
            data[key] = value
    
    return data

#Function to get the data into an xlsx file
def merge_and_append_to_excel(file1_path, file2_path, output_excel_path):
    data1 = read_properties(file1_path)
    data2 = read_properties(file2_path)
    
    #Sets the headers for the xlsx file
    df1 = pd.DataFrame(list(data1.items()), columns=['Key', 'File1']) #File1 can be changed to a diff. name
    df2 = pd.DataFrame(list(data2.items()), columns=['Key', 'File2']) #File2 can be changed to a diff. name
    
    # Load existing Excel file with openpyxl
    try:
        book = load_workbook(output_excel_path)
    except FileNotFoundError:
        book = None

    # Write df1 and df2 to separate sheets
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
        if book:
            writer.book = book  # Use the existing workbook
                                
                            #Provide sheet names as desired
        df1.to_excel(writer, sheet_name="File1", index=False) #File1
        df2.to_excel(writer, sheet_name="File2", index=False) #File2
        
        #Mergas on the left file (df1)
        merged_df = pd.merge(df1, df2, on='Key', how='left', suffixes=('_df1', '_df2'))
        merged_df.to_excel(writer, sheet_name='comparison', index=False) #comparison can be changed to a diff. name

#runs the above programes on the below files. the 'r' before the path allows there to be \ without issues.    
if __name__ == "__main__":
    file1_path = r"C:\Users\stephanie.pietz\Documents\Projects\ORMCO_2401_P0007\02\localization_es.properties" #Include full path & file name
    file2_path = r"C:\Users\stephanie.pietz\Documents\Projects\ORMCO_2401_P0007\02\localization_en.properties"
    output_excel_path = r"C:\Users\stephanie.pietz\Documents\Projects\ORMCO_2401_P0007\02\ES_output.xlsx"

    #Calls the function.
    merge_and_append_to_excel(file1_path, file2_path, output_excel_path)





