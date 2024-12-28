import pandas as pd

# Define the file path
file_path = "TERMWISE/Testwiseanalysisreport_final.xlsx"

# Load the Excel file without any header
df = pd.read_excel(file_path, header=None)

# Step 1: Drop the first 10 rows
df = df.iloc[10:].reset_index(drop=True)

# Step 2: Drop the first row containing 'Bachupally' and reset index
df = df.drop(0).reset_index(drop=True)

# Step 3: Keep student details headers unchanged
# Extract columns for student details
student_details_headers = df.iloc[0, :8].tolist()

# Step 4: Process subject assessment headers and rename
# Create the subject prefix list from the column headers, starting from the 9th column
subjects = []
current_subject = ""

# Loop through the headers to build the subject names with short form
for col_index in range(8, len(df.columns)):
    header = str(df.iloc[1, col_index])
    if "PT" in header or "HF" in header or "NB" in header or "SE" in header:
        # Add the prefix for the subject's first three letters
        subject_short = current_subject[:3] if len(current_subject) >= 3 else current_subject
        if subject_short:
            # Example: English_PT1 ➔ Eng_PT1
            renamed_header = f"{subject_short}_{header}"
            df.iloc[1, col_index] = renamed_header
    else:
        # New subject name found, update current subject
        current_subject = header

# Step 5: Replace the 2nd row with renamed headers and move it to the 3rd row
df.iloc[2] = df.iloc[1]
df = df.drop(1).reset_index(drop=True)

# Combine student details headers with the processed subject headers
final_headers = student_details_headers + df.iloc[1, 8:].tolist()
df.columns = final_headers

# Step 6: Save the processed DataFrame to a new file
output_path = "TERMWISE/Updated_Testwiseanalysisreport_final.xlsx"
df.to_excel(output_path, index=False)

print(f"Processed file saved at: {output_path}")
