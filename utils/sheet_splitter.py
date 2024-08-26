import os
import pandas as pd

# Set the input file path and output folder
input_file = r"C:\Users\User\Downloads\All UDSM Units.xlsx"
output_folder = "colleges"

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Read the Excel file
excel_file = pd.ExcelFile(input_file)

# Get all sheet names
sheet_names = excel_file.sheet_names

# Iterate through sheets, skipping the first one
for sheet_name in sheet_names[1:]:
    # Read the sheet
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    
    # Create the output file path
    output_file = os.path.join(output_folder, f"{sheet_name}.xlsx")
    
    # Save the sheet as a new Excel file
    df.to_excel(output_file, index=False)
    
    print(f"Saved {sheet_name} to {output_file}")

print("All sheets have been saved as separate files.")