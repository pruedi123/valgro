import pandas as pd
import os

def load_data_frame(year):
    # Construct the file path including the 'data_periods' folder
    file_name = f'output_year_{year}.xlsx'
    file_path = os.path.join('data_periods', file_name)  # Cross-platform compatibility
    
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        return None

# When a year is selected (e.g., from a slider)
selected_year = 41
df = load_data_frame(selected_year)

if df is not None:
    # Proceed with using df
    pass
else:
    print(f"Data for year {selected_year} is not available.")

print(df.head())