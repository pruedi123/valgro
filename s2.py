import pandas as pd
from pandas import ExcelWriter

def run_mult_period_factors():
    # Load the Excel file with the portfolio factors
    file_path = 'all_portfolio_annual_factor_20_bps.xlsx'
    df = pd.read_excel(file_path)

    # Create an Excel writer to save multiple sheets in one workbook
    output_file_path = 'row_product_portfolio_annual_factors_years_1_to_41.xlsx'
    writer = ExcelWriter(output_file_path, engine='xlsxwriter')

    # Initialize the row count for Year_2
    year_2_rows = 0

    # Generate Year_2 to Year_41 first
    year_data = []
    for year in range(2, 42):
        products = []
        offset = (year - 2) * 12  # Adjust the offset to ensure correct year calculation

        for i in range(len(df) - offset):
            if year == 2:
                # Year 2 is just the single row value
                product_row = df.iloc[i, 2:]
            else:
                # Start with the product of the first two rows for year 3
                product_row = df.iloc[i, 2:] * df.iloc[i + 12, 2:]
                # For year 4 and beyond, multiply additional rows
                for j in range(2, year - 1):  # Start from year 4's third value
                    product_row *= df.iloc[i + j * 12, 2:]

            products.append(product_row)
        
        # Convert the list of products to a DataFrame and rename columns
        products_df = pd.DataFrame(products, columns=df.columns[2:])
        products_df = products_df.rename(columns=lambda col: f"{col}_yr{year}")
        
        # Append to year data list
        year_data.append(products_df)
        
        # Capture the row count from Year_2
        if year == 2:
            year_2_rows = len(products_df)

    # Now that Year_2 has been calculated, create Year_1
    year_1_data = pd.DataFrame(1, index=range(year_2_rows), columns=df.columns[2:])
    year_1_data = year_1_data.rename(columns=lambda col: f"{col}_yr1")

    # Concatenate Year_1 with Year_2 to Year_41 data
    all_years_data = pd.concat([year_1_data] + year_data, axis=1)

    # Write the combined data to a new sheet in the workbook
    all_years_data.to_excel(writer, sheet_name='Year_1_to_41', index=False)

    # Close the workbook to save it
    writer.close()

    print(f"Products for years 1 to 41 saved in: {output_file_path}")

# Uncomment the line below if you want the module to still be runnable as a standalone script
if __name__ == "__main__":
    run_mult_period_factors()