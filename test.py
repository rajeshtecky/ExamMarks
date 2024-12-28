import pandas as pd

# Load the Excel file without assuming any headers
file_path = 'class8test.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path, header=None)

# List of required columns
required_columns = ['S.No', 'Student Name', 'Roll No', 'Enrollment Code', 'Class - Section', 'Admission No',
                    'English', 'Telugu', 'Hindi', 'Mathematics', 'Science', 'Social', 'ArtificialIntelligence']

# Identify the header row
header_row_index = None
for i, row in df.iterrows():
    if set(required_columns).issubset(set(row)):
        header_row_index = i
        break

if header_row_index is not None:
    print(f"Header row found at index {header_row_index}. Reading data...")
    
    # Read the Excel file again using the identified header row
    df = pd.read_excel(file_path, header=header_row_index)

    # Convert subject columns to numeric, coerce errors to NaN
    subject_columns = required_columns[6:]
    df[subject_columns] = df[subject_columns].apply(pd.to_numeric, errors='coerce')

    # Calculate total marks
    df['Total Marks'] = df[subject_columns].sum(axis=1)

    # Rank students within each section
    df['Rank'] = df.groupby('Class - Section')['Total Marks'].rank(ascending=False, method='min')

    # Save the output files section-wise
    sections = df['Class - Section'].unique()
    for section in sections:
        section_df = df[df['Class - Section'] == section]
        section_df.to_excel(f'section_{section}_ranked.xlsx', index=False)

    print("Files saved successfully.")
else:
    print("The required columns were not found in the file.")
