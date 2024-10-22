import pandas as pd

# Specify the number of years you want to access
num_years = 35  # Change this value as needed

# Define the file path and the worksheet names
file_path = f"/Users/paulruedi/Desktop/py_test2/data_periods/output_year_{num_years}.xlsx"
weighted_allocations_sheet = "Weighted Allocations"  # Corrected worksheet name
transposed_ending_values_sheet = "Transposed Ending Values"

# Read the specified worksheets into DataFrames
weighted_allocations_df = pd.read_excel(file_path, sheet_name=weighted_allocations_sheet)
transposed_ending_values_df = pd.read_excel(file_path, sheet_name=transposed_ending_values_sheet)

# Calculate the average row values for the transposed_ending_values_df
average_row_values = transposed_ending_values_df.mean(axis=1)

# Get the lowest average row value
lowest_average_value = average_row_values.min()

# Get the median average row value
median_average_value = average_row_values.median()

# Display the results
print("Lowest Average Row Value:", lowest_average_value)
print("Median Average Row Value:", median_average_value)