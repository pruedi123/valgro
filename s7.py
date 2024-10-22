import pandas as pd

# Load your data
portfolio_df = pd.read_excel('all_portfolio_annual_factor_20_bps.xlsx', sheet_name='allocation_factors')
cost_df = pd.read_excel('min_values_across_years_and_matrix.xlsx', sheet_name='cost_factors_with_rules')

# Function to extract equity percentage
def extract_equity_percentage(allocation):
    return int(allocation.split(' ')[-1][:-1]) / 100.0 if 'E' in allocation else 0.0

# Step 1: Get cost and allocation for Year_3
year_3_cost = cost_df.loc[1, 'Year_3']  # Fetching Year_3 cost
initial_allocation = 0.10  # Initial allocation as per instructions

# Step 2: Calculate the Year 1 and Year 2 product
# Fetch the factor for initial allocation (LBM 10E) for Year 1 (row 0)
allocation_column = 'LBM 10E'  # Since the allocation is 0.10, equivalent to 10% equity
factor_year_1 = portfolio_df.loc[0, allocation_column]

# Year 2 value is the product of cost and the factor from Year 1
year_2_value = year_3_cost * factor_year_1

# Step 3: Filter for any allocation change
# For simplicity, we assume there is no allocation change (this can be extended with matrix filtering logic)
# You can adjust the logic here to filter based on the allocation change

# Step 4: Calculate the Year 3 ending value
# Fetch the factor for Year 3 using the same allocation (or updated allocation if it changes)
factor_year_3 = portfolio_df.loc[12, allocation_column]  # Move 12 rows down for Year 3 factor
year_3_value = year_2_value * factor_year_3

# Output the results
print(f"Year 3 Cost: {year_3_cost}")
print(f"Year 1 Factor: {factor_year_1}")
print(f"Year 2 Value: {year_2_value}")
print(f"Year 3 Factor: {factor_year_3}")
print(f"Year 3 Ending Value: {year_3_value}")