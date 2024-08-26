import os
import pandas as pd
import glob

# Set the folder path
folder_path = 'college_data'

# Get all CSV files in the folder
csv_files = glob.glob(os.path.join(folder_path, '*.csv'))

# Initialize an empty list to store dataframes
dfs = []

# Process each CSV file
for file in csv_files:
    # Read the CSV file, treating 'N/A' and empty fields as NaN
    df = pd.read_csv(file, na_values=['N/A', ''], keep_default_na=True)
    
    # Extract the college name from the file name
    college_name = os.path.basename(file).split('_')[2].split('.')[0]
    
    # Add the COLLEGE column with the extracted name
    df['COLLEGE'] = college_name
    
    # Append the dataframe to the list
    dfs.append(df)

# Combine all dataframes
combined_df = pd.concat(dfs, ignore_index=True)

# Save the combined dataframe to a new CSV file
# Use the 'na_rep' parameter to represent NaN values as 'N/A' in the output
combined_df.to_csv('combined_college_data.csv', index=False, na_rep='N/A')

print("All CSV files have been combined with the COLLEGE column added.")