import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Set page configuration
st.set_page_config(page_title="CAGR Difference Analysis", layout="wide")

# Streamlit title and description
st.title("CAGR Difference Analysis: Growth vs Value")
st.write("Analysis of performance differences between Growth and Value stocks across different time periods")

# Sidebar controls
st.sidebar.header("Chart Controls")

# Radio button for line display
display_option = st.sidebar.radio(
    "Select Display Option",
    ["CAGR Difference Only", "Moving Average Only", "Both Lines"]
)

# Slider for moving average period
ma_period = st.sidebar.slider(
    "Moving Average Period",
    min_value=5,
    max_value=100,
    value=21,
    step=1
)

# Set display flags based on radio selection
show_cagr_line = display_option in ["CAGR Difference Only", "Both Lines"]
show_ma_line = display_option in ["Moving Average Only", "Both Lines"]

# Load the dataset
@st.cache_data
def load_data():
    file_path = 'val_gro.xlsx'
    df = pd.read_excel(file_path)
    df['Date'] = pd.to_datetime(df['Date'])
    df['Cumulative Growth'] = (1 + df['Growth']).cumprod()
    df['Cumulative Value'] = (1 + df['Value']).cumprod()
    return df

data_df = load_data()

def calculate_cagr_differences(data, period_years):
    period_length_months = period_years * 12
    cagr_differences = []

    for start_index in range(len(data) - period_length_months + 1):
        period_df = data.iloc[start_index:start_index + period_length_months]

        growth_start = period_df['Cumulative Growth'].iloc[0]
        growth_end = period_df['Cumulative Growth'].iloc[-1]
        value_start = period_df['Cumulative Value'].iloc[0]
        value_end = period_df['Cumulative Value'].iloc[-1]

        growth_cagr = ((growth_end / growth_start) ** (1 / period_years)) - 1
        value_cagr = ((value_end / value_start) ** (1 / period_years)) - 1

        cagr_difference = growth_cagr - value_cagr

        cagr_differences.append({
            'Start Date': period_df['Date'].iloc[0],
            'End Date': period_df['Date'].iloc[-1],
            'CAGR Difference': cagr_difference
        })

    return pd.DataFrame(cagr_differences)

# Calculate CAGR differences for different periods
periods = [5, 10, 15, 20, 25, 30]
cagr_dfs = {
    period: calculate_cagr_differences(data_df, period)
    for period in periods
}

# Calculate moving averages
for df in cagr_dfs.values():
    df['MA'] = df['CAGR Difference'].rolling(window=ma_period).mean()

# Create the plots with more vertical space between subplots
fig, axs = plt.subplots(6, 1, figsize=(14, 30))  # Increased height from 24 to 30
plt.subplots_adjust(hspace=0.8)  # Increased from 0.4 to 0.8

def add_explanation_text(ax):
    ax.text(0.02, 0.95, 'If trending up: Growth outperforming Value', 
            transform=ax.transAxes, fontsize=8, verticalalignment='top')
    ax.text(0.02, 0.90, 'If Trending Down: Value outperforming Growth', 
            transform=ax.transAxes, fontsize=8, verticalalignment='top')

def plot_subplot(ax, data, period, color):
    if show_cagr_line:
        ax.plot(data['End Date'], data['CAGR Difference'], label=f'{period}-Year Period', color=color)
    if show_ma_line:
        ax.plot(data['End Date'], data['MA'], 
                label=f'{ma_period}-Period MA', 
                color='black', 
                linestyle='--')
    
    # Format dates for x-axis
    years = mdates.YearLocator(5)
    years_fmt = mdates.DateFormatter('%Y')
    ax.xaxis.set_major_locator(years)
    ax.xaxis.set_major_formatter(years_fmt)
    
    # Set title with more padding
    ax.set_title(f'CAGR Difference (Growth - Value) for {period}-Year Periods', pad=20)
    
    # Add date range below the chart with adjusted position
    start_date = data['Start Date'].iloc[0].strftime('%Y')
    end_date = data['End Date'].iloc[-1].strftime('%Y')
    ax.text(0.5, -0.25, f'Date Range: {start_date} - {end_date}', 
            horizontalalignment='center',
            transform=ax.transAxes,
            fontsize=10)
    
    ax.set_ylabel('CAGR Difference')
    ax.axhline(0, color='red', linewidth=0.8, linestyle='--')
    ax.grid(True)
    ax.legend()
    add_explanation_text(ax)
    
    # Rotate and align the tick labels so they look better
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right')

# Colors for different periods
colors = ['blue', 'orange', 'green', 'red', 'purple', 'brown']

# Plot each subplot
for i, (period, color) in enumerate(zip(periods, colors)):
    plot_subplot(axs[i], cagr_dfs[period], period, color)

# Set x-label for the bottom subplot
axs[-1].set_xlabel('End Date of Period')

# Adjust layout to prevent label cutoff with more padding
plt.tight_layout(rect=[0, 0.03, 1, 0.95], h_pad=2.0)  # Added h_pad parameter

# Display the plot in Streamlit
st.pyplot(fig)

# Add explanatory text below the chart
st.markdown("""
### How to Interpret the Charts
- **Trend Direction**: 
    - Upward trend indicates Growth outperforming Value
    - Downward trend indicates Value outperforming Growth
- **Zero Line**: The red dashed line represents zero difference in performance
- **Moving Average**: Helps smooth out short-term fluctuations to show longer-term trends
""")

# Display current settings
st.sidebar.markdown("""
### Current Settings
""")
st.sidebar.write(f"Display Mode: {display_option}")
st.sidebar.write(f"Moving Average Period: {ma_period}")