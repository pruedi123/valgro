import os
import pandas as pd
import numpy as np

def combine_transposed_ending_values(data_folder, max_years=41):
    """Combines the last columns from 'Transposed Ending Values' across multiple files."""
    
    # Initialize a list to store each year's last column data
    all_columns = []

    # Loop through years 1 to max_years (inclusive)
    for year in range(1, max_years + 1):
        # Construct the full file path for the current year
        file_name = os.path.join(data_folder, f'data_{year}_year_periods.xlsx')
        sheet_name = 'Transposed Ending Values'
        
        # Read the 'Transposed Ending Values' sheet from the file
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        
        # Extract the last column
        last_column = df.iloc[:, -1].copy()  # Get the last column of the DataFrame
        
        # Add the column to the list with the proper column name
        all_columns.append(last_column.rename(f'Year_{year}'))

    # Concatenate all the columns into a single DataFrame
    final_df = pd.concat(all_columns, axis=1)

    # Save the final DataFrame to an Excel file
    output_file = os.path.join(data_folder, 'combined_transposed_ending_values.xlsx')  # Save to the same folder
    final_df.to_excel(output_file, index=False)

    print(f"Data saved to {output_file}")

# Uncomment the line below if you want the module to still be runnable as a standalone script
# if __name__ == "__main__":
    data_folder = '/Users/paulruedi/Desktop/py_test2/data_periods'
    combine_transposed_ending_values(data_folder)