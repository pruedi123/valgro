import pandas as pd
import numpy as np

def run_cost_matrix():
    # Load the Excel file
    file_path = 'min_values_across_years_and_matrix.xlsx'

    # Load the matrix sheet
    matrix_df = pd.read_excel(file_path, sheet_name='Matrix')

    # Initialize lists to store the results
    allocation_costs = []
    lowest_values = []

    # Define a dictionary for the allocation costs
    allocation_cost_dict = {
        'LBM 100E': 1.0,
        'LBM 90E': 0.90,
        'LBM 80E': 0.80,
        'LBM 70E': 0.70,
        'LBM 60E': 0.60,
        'LBM 50E': 0.50,
        'LBM 40E': 0.40,
        'LBM 30E': 0.30,
        'LBM 20E': 0.20,
        'LBM 10E': 0.10,
        'LBM 100F': 0.0,
    }

    # Parameter to apply or not apply the rule
    apply_rule = True  # Set this to False if you don't want to apply the rule

    # Loop through the years 1 to 41 (starting with 1 for year 1)
    previous_cost = None
    previous_allocation = None

    for year in range(1, 42):
        year_col = f'Year_{year}'

        if year == 1:
            allocation_costs.append(0)
            lowest_values.append(1)
            previous_cost = 0
            previous_allocation = 'LBM 100F'
        else:
            # Extract the column for the current year
            current_year_values = matrix_df[year_col]
            
            # Find the index of the minimum value in the current year
            min_index = current_year_values.idxmin()
            
            # Get the corresponding allocation and its value
            allocation = matrix_df.at[min_index, 'Allocation']
            value = current_year_values[min_index]
            
            # Get the cost associated with the allocation
            cost = allocation_cost_dict.get(allocation, np.nan)
            
            # Apply the rule if specified
            if apply_rule and cost < previous_cost:
                cost = previous_cost
                allocation = previous_allocation  # Use the previous year's allocation instead

            # Append the allocation cost
            allocation_costs.append(cost)
            previous_allocation = allocation

            # Debugging information
            # print(f"Year {year}: Allocation = {allocation}, Cost = {cost}")

            # Now correctly filter the Matrix tab using the selected allocation for this year
            filtered_values = matrix_df.loc[matrix_df['Allocation'] == allocation, year_col]
            
            # Check if any values were found, and handle cases where no matching allocation is found
            if not filtered_values.empty:
                filtered_value = filtered_values.values[0]
            else:
                filtered_value = np.nan  # Or use 0 or another default value if preferred
                print(f"Warning: No matching value found for allocation {allocation} in Year {year}")

            lowest_values.append(filtered_value)
            
            # Update previous cost and value
            previous_cost = cost

    # Determine the sheet name based on whether the rule was applied
    sheet_name = 'cost_factors_with_rules' if apply_rule else 'cost_factors'

    # Create the DataFrame for the cost_factors tab
    cost_factors_df = pd.DataFrame([allocation_costs, lowest_values], columns=[f'Year_{year}' for year in range(1, 42)])
    cost_factors_df.insert(0, 'Factor', ['Lowest Cost Allocation', 'Lowest Cost Value'])

    # Reverse the order of the years' columns for the reverse_matrix
    reversed_columns = ['Allocation'] + [f'Year_{year}' for year in range(41, 0, -1)]
    reverse_matrix_df = matrix_df[['Allocation'] + [f'Year_{year}' for year in range(1, 42)]].copy()
    reverse_matrix_df = reverse_matrix_df[reversed_columns]

    # Write both the cost_factors and reverse_matrix to the Excel file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        cost_factors_df.to_excel(writer, sheet_name=sheet_name, index=False)
        reverse_matrix_df.to_excel(writer, sheet_name='reverse_matrix', index=False)

    print(f"Cost factors added to {file_path} with {'the no allocation decrease rule applied' if apply_rule else 'no rules applied'}.")
    print("Reverse matrix added to the same file with years in reverse order.")

# Uncomment the line below if you want the module to still be runnable as a standalone script
if __name__ == "__main__":
    run_cost_matrix()