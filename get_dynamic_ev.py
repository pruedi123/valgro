import pandas as pd
import numpy as np
from concurrent.futures import ProcessPoolExecutor

import time

start_time = time.time()

# Your existing code that processes the data goes here


def extract_equity_percentage(allocation):
    try:
        if allocation == 'LBM 100F':
            return 0.0
        elif allocation and 'LBM' in allocation:
            number = int(allocation.split(' ')[-1][:-1])
            return number / 100.0
        else:
            return np.nan
    except Exception as e:
        return np.nan

def get_allocation_for_year(allocation_df, year_column):
    return allocation_df.loc[allocation_df['Factor'] == 'Lowest Cost Allocation', year_column].values[0]

def determine_initial_allocation(allocation_df, total_years):
    year_column = f'Year_{total_years}'
    allocation_percentage = get_allocation_for_year(allocation_df, year_column)
    allocation_mapping = {
        1: 'LBM 100E', 0.9: 'LBM 90E', 0.8: 'LBM 80E', 0.7: 'LBM 70E',
        0.6: 'LBM 60E', 0.5: 'LBM 50E', 0.4: 'LBM 40E', 0.3: 'LBM 30E',
        0.2: 'LBM 20E', 0.1: 'LBM 10E', 0: 'LBM 100F'
    }
    rounded_value = round(allocation_percentage, 1)
    return allocation_mapping.get(rounded_value, 'Unknown Allocation')

def fetch_start_value(allocation_df, year_column):
    return allocation_df.loc[1, year_column]

def calculate_allocations_and_values(portfolio_df, matrix_df, allocation_df, time_period, start_index):
    start_allocation = determine_initial_allocation(allocation_df, time_period)
    if start_allocation == 'Unknown Allocation':
        raise ValueError("The initial allocation could not be determined.")
    
    year_column = f'Year_{time_period}'
    start_value = fetch_start_value(allocation_df, year_column)
    years = [1]
    allocations = ['NA']
    values = [start_value]
    factors_used = [start_value]
    detailed_years_data = []

    detailed_years_data.append({
        'Year': 1,
        'Allocation': 'NA',
        'Factor Used': 1,
        'Ending Value': start_value
    })

    row_to_use_year_2 = start_index
    factor_for_year_2 = portfolio_df[start_allocation].iloc[row_to_use_year_2]
    current_value = values[-1] * factor_for_year_2
    years.append(2)
    allocations.append(start_allocation)
    factors_used.append(factor_for_year_2)
    values.append(current_value)

    detailed_years_data.append({
        'Year': 2,
        'Allocation': start_allocation,
        'Factor Used': factor_for_year_2,
        'Ending Value': current_value
    })

    for year in range(3, time_period + 1):
        matrix_year = time_period - (year - 1) + 1
        matrix_column = f'Year_{matrix_year}'
        filter_value = values[-1]

        filtered_allocation = matrix_df[
            (matrix_df[matrix_column] <= filter_value) &
            (matrix_df['Allocation'].apply(extract_equity_percentage) < extract_equity_percentage(start_allocation))
        ].copy()

        if not filtered_allocation.empty:
            filtered_allocation['Equity_Percentage'] = filtered_allocation['Allocation'].apply(extract_equity_percentage)
            filtered_allocation = filtered_allocation.sort_values(by='Equity_Percentage')
            start_allocation = filtered_allocation.iloc[0]['Allocation']

        row_to_use = (year - 2) * 12 + start_index
        factor_for_current_year = portfolio_df[start_allocation].iloc[row_to_use]
        years.append(year)
        allocations.append(start_allocation)
        factors_used.append(factor_for_current_year)
        current_value = values[-1] * factor_for_current_year
        values.append(current_value)

        detailed_years_data.append({
            'Year': year,
            'Allocation': start_allocation,
            'Factor Used': factor_for_current_year,
            'Ending Value': current_value
        })

    df_results = pd.DataFrame({
        'Year': years,
        'Allocation': allocations,
        'Factor Used': factors_used,
        'Ending Value': values
    })

    return df_results, pd.DataFrame(detailed_years_data)

def process_all_rows_for_start_year(portfolio_df, matrix_df, allocation_df, time_period, start_row, end_row):
    all_values = []
    all_allocations = []
    detailed_data_list = []
    for i in range(start_row, end_row + 1):
        results, detailed_data = calculate_allocations_and_values(portfolio_df, matrix_df, allocation_df, time_period, i)
        all_values.append(results['Ending Value'].values)
        all_allocations.append(results['Allocation'].values)
        detailed_data['Run'] = i
        detailed_data_list.append(detailed_data)

    df_values = pd.DataFrame(all_values, columns=[f'Year_{i}' for i in range(1, time_period + 1)])
    df_allocations = pd.DataFrame(all_allocations, columns=[f'Year_{i}' for i in range(1, time_period + 1)])
    
    if detailed_data_list:
        detailed_data_df = pd.concat(detailed_data_list, ignore_index=True)
    else:
        detailed_data_df = pd.DataFrame()

    return df_values, df_allocations, detailed_data_df

def run_simulation_across_years(portfolio_df, matrix_df, allocation_df, start_row, end_row, max_years):
    all_values_results = []
    all_allocations_results = []
    all_detailed_data = []

    for start_year in range(2, max_years + 1):
        df_values, df_allocations, detailed_data_df = process_all_rows_for_start_year(
            portfolio_df, matrix_df, allocation_df, start_year, start_row, end_row)

        num_padding = max_years - len(df_values.columns)
        for row_idx in range(df_values.shape[0]):
            values_with_padding = df_values.iloc[row_idx].tolist() + [0] * num_padding
            allocations_with_padding = df_allocations.iloc[row_idx].tolist() + ['NA'] * num_padding
            all_values_results.append(values_with_padding[:max_years])
            all_allocations_results.append(allocations_with_padding[:max_years])
        
        all_detailed_data.append(detailed_data_df)

    column_names = [f'Year_{i}' for i in range(1, max_years + 1)]
    all_values_df = pd.DataFrame(all_values_results, columns=column_names)
    all_allocations_df = pd.DataFrame(all_allocations_results, columns=column_names)

    if all_detailed_data:
        all_detailed_data_df = pd.concat(all_detailed_data, ignore_index=True)
    else:
        all_detailed_data_df = pd.DataFrame()

    return all_values_df, all_allocations_df, all_detailed_data_df

def calculate_weighted_allocation(values_df, allocations_df):
    weighted_allocations = []
    for year in values_df.columns:
        allocations = allocations_df[year]
        values = values_df[year]
        
        # Filter non-zero values
        non_zero_mask = values != 0
        non_zero_values = values[non_zero_mask]
        non_zero_allocations = allocations[non_zero_mask].apply(extract_equity_percentage)
        
        # Calculate weighted average if there are non-zero values
        weighted_avg = np.average(non_zero_allocations, weights=non_zero_values) if len(non_zero_values) > 0 else 0
        weighted_allocations.append(weighted_avg)

    weighted_allocations[0] = 0  # Year 1 allocation typically set to 0
    return pd.DataFrame([weighted_allocations], columns=values_df.columns)
    print(weighted_allocations_df)

def get_last_non_zero_values(values_df):
    return values_df.apply(lambda row: row[row != 0].iloc[-1] if any(row != 0) else np.nan, axis=1)

def accumulate_results_for_rows(portfolio_df, matrix_df, allocation_df, max_years, start_rows):
    all_portfolio_values = []
    all_weighted_allocations = []
    all_transposed_ending_values = []
    all_detailed_data = []

    for start_row in start_rows:
        all_values_df, all_allocations_df, detailed_data_df = run_simulation_across_years(
            portfolio_df, matrix_df, allocation_df, start_row, start_row, max_years)

        # Drop the last 12 rows from the Year 1 output
        if len(all_values_df) > 12:  # Check if there are enough rows to drop
            all_values_df = all_values_df[:-12]  # Drop the last 12 rows

        portfolio_val = pd.DataFrame(all_values_df.sum(), columns=[f'Total_{start_row}']).T
        all_portfolio_values.append(portfolio_val)

        weighted_allocations_df = calculate_weighted_allocation(all_values_df, all_allocations_df)
        weighted_allocations_df.index = [f'Run_{start_row}']
        all_weighted_allocations.append(weighted_allocations_df)

        last_non_zero_values_df = pd.DataFrame(get_last_non_zero_values(all_values_df)).T
        last_non_zero_values_df.index = [f'Run_{start_row}']
        all_transposed_ending_values.append(last_non_zero_values_df)

        all_detailed_data.append(detailed_data_df)

    portfolio_values_df = pd.concat(all_portfolio_values, ignore_index=True)
    weighted_allocations_df = pd.concat(all_weighted_allocations)
    transposed_ending_values_df = pd.concat(all_transposed_ending_values)
    detailed_data_df = pd.concat(all_detailed_data, ignore_index=True)

    return portfolio_values_df, weighted_allocations_df, transposed_ending_values_df, detailed_data_df
def add_column_of_ones(df):
    df.insert(0, 'Year_1', 1)
    return df

def rename_columns_to_match(df, reference_df):
    df.columns = reference_df.columns
    return df

# Load the data from Excel
portfolio_df = pd.read_excel('all_portfolio_annual_factor_20_bps.xlsx', sheet_name='allocation_factors')
matrix_df = pd.read_excel('dynamic_data.xlsx', sheet_name='matrix')
allocation_df = pd.read_excel('min_values_across_years_and_matrix.xlsx', sheet_name='cost_factors_with_rules')

# Parameters for the simulation
max_years_range = range(1, 6)  # Loop from 1 to 30
act_start_row = 1  # 1-based index for starting row
output_directory = '/Users/paulruedi/Desktop/py_test2/data_periods/'

# Loop through max_years values
for max_years in max_years_range:
    if max_years == 1:
        num_rows_to_process = 1152 - (max_years * 12) + 12  # Adjust for Year 1
    else:
        num_rows_to_process = 1152 - (max_years * 12) + 24  # For other years

    start_rows = list(range(act_start_row - 1, act_start_row - 1 + num_rows_to_process))

    portfolio_values_df, weighted_allocations_df, transposed_ending_values_df, detailed_data_df = accumulate_results_for_rows(
        portfolio_df, matrix_df, allocation_df, max_years, start_rows)
    if 'Year_1' in portfolio_values_df.columns:
        portfolio_values_df['Year_1'] += 1

    transposed_ending_values_df = add_column_of_ones(transposed_ending_values_df)
    transposed_ending_values_df = rename_columns_to_match(transposed_ending_values_df, portfolio_values_df)

    output_filename = f'{output_directory}output_year_{max_years}.xlsx'
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        portfolio_values_df.to_excel(writer, sheet_name='Portfolio Values', index=True)
        weighted_allocations_df.to_excel(writer, sheet_name='Weighted Allocations', index=True)
        transposed_ending_values_df.to_excel(writer, sheet_name='Transposed Ending Values', index=True)
        detailed_data_df.to_excel(writer, sheet_name='Detailed Data', index=False)

    print(f"DataFrames for max_years {max_years} saved to {output_filename}.")

    end_time = time.time()
elapsed_time = end_time - start_time
print(f"Execution time: {elapsed_time:.2f} seconds")
