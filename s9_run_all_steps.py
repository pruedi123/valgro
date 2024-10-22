
import pandas as pd


import s1_get_factors_all_allocations
import s2_get_mult_period_factors
import s3_get_mins_and_matrix
import s4_get_cost_matrix
import s5_get_11_alloc_all_years_factors
import s6_get_each_row_end_value
import s7_get_all_years_ending_values
import s8_get_final_ending_data_all_years
import s9_run_all_steps  # Replace with actual module name that has the combine_transposed_ending_values function

def run_all_steps():
    """Runs all steps from s1 to s8 and combines transposed ending values."""
    
    # Step 1: Run factors for all allocations
    s1_get_factors_all_allocations.run_factors_all_allocations()

    # Step 2: Run multi-period factors
    s2_get_mult_period_factors.run_mult_period_factors()

    # Step 3: Run mins and matrix
    s3_get_mins_and_matrix.run_mins_and_matrix()

    # Step 4: Run cost matrix
    s4_get_cost_matrix.run_cost_matrix()

    # Step 5: Run allocation factors for all years
    s5_get_11_alloc_all_years_factors.run_alloc_all_years_factors()

    # Step 6: Run each row end value
    s6_get_each_row_end_value.run_each_row_end_value()

    # Step 7: Run all years ending values
    portfolio_df = pd.read_excel('all_portfolio_annual_factor_20_bps.xlsx', sheet_name='allocation_factors')
    matrix_df = pd.read_excel('dynamic_data.xlsx', sheet_name='matrix')
    allocation_df = pd.read_excel('min_values_across_years_and_matrix.xlsx', sheet_name='cost_factors_with_rules')
    output_folder = '/Users/paulruedi/Desktop/py_test2/data_periods'
    
    s7_get_all_years_ending_values.run_all_years_ending_values(portfolio_df, matrix_df, allocation_df, output_folder)

    # Step 8: Run final ending data for all years
    s8_get_final_ending_data_all_years.run_final_ending_data_all_years()

    # Combine Transposed Ending Values
    your_module_name.combine_transposed_ending_values(output_folder)

if __name__ == "__main__":
    run_all_steps()