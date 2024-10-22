import pandas as pd

def run_mins_and_matrix():
    # Load the Excel file
    file_path = 'row_product_portfolio_annual_factors_years_1_to_41.xlsx'

    # Load the worksheet named 'Year_1_to_41'
    df = pd.read_excel(file_path, sheet_name='Year_1_to_41')

    # Initialize a list to store each allocation's data
    data = []

    # Extract the unique years from the column names
    years = set(col.split('_yr')[1] for col in df.columns)

    # Loop through the years by extracting relevant columns
    for year in sorted(years, key=int):
        # Filter columns corresponding to the current year
        year_columns = [col for col in df.columns if col.endswith(f'_yr{year}')]

        # Calculate the minimum value for each allocation
        min_values = df[year_columns].min()

        # If it's the first year, initialize the data structure with allocation names
        if year == '1':  # The first year
            data = [[alloc.split('_yr')[0]] for alloc in year_columns]

        # Append the minimum values for the current year to each allocation's row
        for i, value in enumerate(min_values):
            data[i].append(value)

    # Convert the list of lists to a DataFrame for minimum values
    columns = ['Allocation'] + [f'Year_{year}' for year in sorted(years, key=int)]
    min_values_df = pd.DataFrame(data, columns=columns)

    # Calculate the inverse values (1/min_value) for the matrix
    matrix_df = min_values_df.copy()
    for col in matrix_df.columns[1:]:  # Skip the 'Allocation' column
        matrix_df[col] = 1 / matrix_df[col]

    # Save the result to a new Excel file with two sheets
    output_file_path = 'min_values_across_years_and_matrix.xlsx'
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        min_values_df.to_excel(writer, sheet_name='Min Values', index=False)
        matrix_df.to_excel(writer, sheet_name='Matrix', index=False)

    print(f"Minimum values and inverse matrix saved in: {output_file_path}")

# Uncomment the line below if you want the module to still be runnable as a standalone script
if __name__ == "__main__":
    run_mins_and_matrix()