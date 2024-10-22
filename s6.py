import pandas as pd
import numpy as np

# Load the Excel files globally
allocation_file_path = 'all_portfolio_annual_factor_20_bps.xlsx'
cost_factors_file_path = 'min_values_across_years_and_matrix.xlsx'

allocation_df = pd.read_excel(allocation_file_path, sheet_name='allocation_factors')
cost_factors_df = pd.read_excel(cost_factors_file_path, sheet_name='cost_factors_with_rules')

def run_each_row_end_value():
    def calculate_ending_values(start_index, num_years):
        # Initialize a list to store ending values for all years
        ending_values = []

        # Loop through each year (1 to num_years)
        for year in range(1, num_years + 1):
            if year == 1:
                ending_values.append(1)  # Year 1 is always set to 1
            else:
                allocation_percentage = cost_factors_df.loc[0, f'Year_{year}']
                if allocation_percentage == 0:
                    allocation_column = 'LBM 100F'
                else:
                    equity_percentage = int(allocation_percentage * 100)
                    allocation_column = f'LBM {equity_percentage}E'

                # Calculate the cumulative product for the current year
                product = 1
                for i in range(year - 1):  # year - 1 because Year 1 is always 1 and we start multiplying from Year 2
                    row_to_access = start_index + i * 12
                    if row_to_access < len(allocation_df):  # Check if the row is within range
                        factor_value = allocation_df.at[row_to_access, allocation_column]
                        if not pd.isna(factor_value):  # Check if the cell is not blank
                            product *= factor_value
                        else:
                            product *= 0  # Multiply by 0 if the value is NaN
                            break
                    else:
                        product *= 0  # Multiply by 0 if the row is out of range
                        break

                # Store the ending value for this year
                ending_values.append(product)

        return ending_values

    # Define the number of years and calculate num_rows
    num_years = 41  # Use all years for calculation
    num_rows = 1152  # Include all rows up to the last one

    # Initialize a list to store all results
    all_results = []

    # Calculate allocations once, as they will be the same for all starting rows
    allocations = []
    for year in range(1, num_years + 1):
        allocation_percentage = cost_factors_df.loc[0, f'Year_{year}']
        if allocation_percentage == 0:
            allocations.append('LBM 100F')
        else:
            equity_percentage = int(allocation_percentage * 100)
            allocations.append(f'LBM {equity_percentage}E')

    # Loop through each index row value from 1 to num_rows
    for start_index in range(num_rows):
        ending_values = calculate_ending_values(start_index, num_years)
        ending_values.insert(0, start_index + 1)  # Add the starting row index at the beginning
        all_results.append(ending_values)

    # Create a DataFrame for all results
    column_names = ['Starting Row'] + [f'Year_{year} ({alloc})' for year, alloc in enumerate(allocations, start=1)]
    all_results_df = pd.DataFrame(all_results, columns=column_names)

    # Calculate normalized values by dividing each column by its minimum non-zero value
    normalized_df = all_results_df.copy()
    for column in normalized_df.columns[1:]:  # Skip 'Starting Row' column
        min_non_zero_value = normalized_df[column].replace(0, np.nan).min()  # Find the minimum non-zero value
        normalized_df[column] = normalized_df[column] / min_non_zero_value

    # Output the final DataFrame and the normalized DataFrame to an Excel file
    output_file_path = 'ending_values_static_all_years.xlsx'
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        all_results_df.to_excel(writer, sheet_name='ending_values', index=False)
        normalized_df.to_excel(writer, sheet_name='ending_values_adjusted', index=False)

    print(f"Results saved to {output_file_path}")

# Uncomment the line below if you want the module to still be runnable as a standalone script
if __name__ == "__main__":
    run_each_row_end_value()

print(cost_factors_df)