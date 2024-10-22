import pandas as pd

def run_factors_all_allocations():
    ##########
    ####### To get all factors for all allocations for all years 1-41, run this module first #########
    ####### Then run get_mult_period_factors.py ########
    ####### Then run get_mins_and_matrix.py ########

    # Load the Excel file
    file_path = 'annual_factors_100E_100F_20_bps.xlsx'
    df = pd.read_excel(file_path)

    # Create new columns for each allocation (90E, 80E, ..., 10E)
    allocations = [90, 80, 70, 60, 50, 40, 30, 20, 10]
    for allocation in allocations:
        equity_weight = allocation / 100.0
        fixed_income_weight = 1 - equity_weight
        column_name = f'LBM {allocation}E'
        df[column_name] = equity_weight * df['LBM 100E'] + fixed_income_weight * df['LBM 100F']

    # Define the desired order of columns
    ordered_columns = ['Start', 'End', 'LBM 100E'] + [f'LBM {allocation}E' for allocation in allocations] + ['LBM 100F']

    # Reorder the DataFrame columns
    df = df[ordered_columns]

    # Save the updated DataFrame back to the Excel file
    output_file_path = 'all_portfolio_annual_factor_20_bps.xlsx'
    df.to_excel(output_file_path, index=False)

    # Create a second DataFrame without 'Start' and 'End', and with an index starting at 1
    df_no_start_end = df.drop(columns=['Start', 'End'])
    df_no_start_end.insert(0, 'Index', range(0, len(df_no_start_end) ))

    # Write both DataFrames to separate tabs in the Excel file
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='All Data', index=False)
        df_no_start_end.to_excel(writer, sheet_name='allocation_factors', index=False)

    print(f"Updated Excel file saved with two tabs: 'All Data' and 'Indexed Data' in: {output_file_path}")

# Uncomment the line below if you want the module to still be runnable as a standalone script
if __name__ == "__main__":
    run_factors_all_allocations()