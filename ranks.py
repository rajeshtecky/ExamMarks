import pandas as pd
import os

# Directory containing the Excel files
root_dir = 'D:/rajesh/examreports/output'
output_file = 'D:/rajesh/examreports/combined_output.xlsx'

all_data = []

# Walk through the directory
for subdir, _, files in os.walk(root_dir):
    for file in files:
        if file.endswith('.xlsx'):
            file_path = os.path.join(subdir, file)
            print(f'Processing file: {file_path}')
            
            # Read the Excel file
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Add a column to identify the source file and sheet
                df['Source File'] = file
                df['Sheet Name'] = sheet_name
                
                all_data.append(df)

# Combine all data into a single DataFrame
combined_df = pd.concat(all_data, ignore_index=True)

# Write the combined DataFrame to a new Excel file
combined_df.to_excel(output_file, index=False)
print(f'Combined data saved to: {output_file}')
