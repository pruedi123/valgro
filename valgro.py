import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Set page configuration
st.set_page_config(page_title="CAGR Difference Analysis", layout="wide")

# Streamlit title
st.title("CAGR Difference Analysis")

# Sidebar controls
st.sidebar.header("Chart Controls")

# Analysis type selector
analysis_type = st.sidebar.selectbox(
    "Select Analysis Type",
    ["Value vs Growth", "Large Growth vs Small Value"]
)

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

# Load the datasets
@st.cache_data
def load_data(analysis_type):
    if analysis_type == "Value vs Growth":
        file_path = 'val_gro.xlsx'
        df = pd.read_excel(file_path)
        df['Date'] = pd.to_datetime(df['Date'])
        df['Cumulative Growth'] = (1 + df['Growth']).cumprod()
        df['Cumulative Value'] = (1 + df['Value']).cumprod()
    else:  # Large Growth vs Small Value
        file_path = 'lgsv.xlsx'
        df = pd.read_excel(file_path)
        df['Date'] = pd.to_datetime(df['Date'])
        df['Cumulative Large Growth'] = (1 + df['Large Growth']).cumprod()
        df['Cumulative Small Value'] = (1 + df['Small Value']).cumprod()
    return df

data_df = load_data(analysis_type)

def calculate_cagr_differences(data, period_years, analysis_type):
    period_length_months = period_years * 12
    cagr_differences = []

    for start_index in range(len(data) - period_length_months + 1):
        period_df = data.iloc[start_index:start_index + period_length_months]

        if analysis_type == "Value vs Growth":
            start_1 = period_df['Cumulative Growth'].iloc[0]
            end_1 = period_df['Cumulative Growth'].iloc[-1]
            start_2 = period_df['Cumulative Value'].iloc[0]
            end_2 = period_df['Cumulative Value'].iloc[-1]
        else:  # Large Growth vs Small Value
            start_1 = period_df['Cumulative Large Growth'].iloc[0]
            end_1 = period_df['Cumulative Large Growth'].iloc[-1]
            start_2 = period_df['Cumulative Small Value'].iloc[0]
            end_2 = period_df['Cumulative Small Value'].iloc[-1]

        cagr_1 = ((end_1 / start_1) ** (1 / period_years)) - 1
        cagr_2 = ((end_2 / start_2) ** (1 / period_years)) - 1

        cagr_difference = cagr_1 - cagr_2

        cagr_differences.append({
            'Start Date': period_df['Date'].iloc[0],
            'End Date': period_df['Date'].iloc[-1],
            'CAGR Difference': cagr_difference
        })

    return pd.DataFrame(cagr_differences)

# Calculate CAGR differences for different periods
periods = [5, 10, 15, 20, 25, 30]
cagr_dfs = {
    period: calculate_cagr_differences(data_df, period, analysis_type)
    for period in periods
}

# Calculate moving averages
for df in cagr_dfs.values():
    df['MA'] = df['CAGR Difference'].rolling(window=ma_period).mean()

# Create the plots
fig, axs = plt.subplots(6, 1, figsize=(14, 30))
plt.subplots_adjust(hspace=0.8)

def add_explanation_text(ax):
    if analysis_type == "Value vs Growth":
        ax.text(0.02, 0.95, 'If trending up: Growth outperforming Value', 
                transform=ax.transAxes, fontsize=8, verticalalignment='top')
        ax.text(0.02, 0.90, 'If Trending Down: Value outperforming Growth', 
                transform=ax.transAxes, fontsize=8, verticalalignment='top')
    else:
        ax.text(0.02, 0.95, 'If trending up: Large Growth outperforming Small Value', 
                transform=ax.transAxes, fontsize=8, verticalalignment='top')
        ax.text(0.02, 0.90, 'If Trending Down: Small Value outperforming Large Growth', 
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
    title = f'CAGR Difference ({analysis_type}) for {period}-Year Periods'
    ax.set_title(title, pad=20)
    
    # Add date range below the chart
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

# Adjust layout
plt.tight_layout(rect=[0, 0.03, 1, 0.95], h_pad=2.0)

# Display the plot in Streamlit
st.pyplot(fig)

# Add explanatory text below the chart
st.markdown("""
### How to Interpret the Charts
- **Trend Direction**: 
    - Upward trend indicates first category outperforming second category
    - Downward trend indicates second category outperforming first category
- **Zero Line**: The red dashed line represents zero difference in performance
- **Moving Average**: Helps smooth out short-term fluctuations to show longer-term trends
""")

# Display current settings
st.sidebar.markdown("""
### Current Settings
""")
st.sidebar.write(f"Analysis Type: {analysis_type}")
st.sidebar.write(f"Display Mode: {display_option}")
st.sidebar.write(f"Moving Average Period: {ma_period}")
