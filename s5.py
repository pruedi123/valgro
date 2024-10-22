import pandas as pd

# Load the Excel file
file_path = 'min_values_across_years_and_matrix.xlsx'

# Load the matrix sheet
matrix_df = pd.read_excel(file_path, sheet_name='Matrix')

# Reverse the order of the years' columns
reversed_columns = ['Allocation'] + [f'Year_{year}' for year in range(41, 0, -1)]
reverse_matrix_df = matrix_df[reversed_columns]

# Write the reversed matrix to a new sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    reverse_matrix_df.to_excel(writer, sheet_name='reverse_matrix', index=False)

print(f"Reverse matrix added to {file_path} with years in reverse order.")