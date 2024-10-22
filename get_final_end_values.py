import pandas as pd
import os

# Define the file path
file_path = '/Users/paulruedi/Desktop/py_test2/data_periods'

# Initialize the new DataFrame with Year 1 data
df_combined = pd.read_excel(os.path.join(file_path, 'output_year_1.xlsx'))

# Loop through years 2 to 41
for year in range(1, 42):
    # Load the current year's data
    df_year = pd.read_excel(os.path.join(file_path, f'output_year_{year}.xlsx'))
    
    # Print the number of rows in the current year's DataFrame
    print(f"Year {year} has {df_year.shape[0]} rows.")

    # Get the last column of the current year's DataFrame
    last_column = df_year.iloc[:, -1]  # Assuming the last column is what you want
    
    # Add the last column to the combined DataFrame
    df_combined[f'Year_{year}'] = last_column.reset_index(drop=True)

# Check the combined DataFrame's shape
print(f"Combined DataFrame shape: {df_combined.shape}")

# Now df_combined contains Year 1 and last columns from Years 2 to 41
df_combined.to_excel(os.path.join(file_path, 'combined_years_output.xlsx'), index=False)

print("DataFrame created and saved as 'combined_years_output.xlsx'")