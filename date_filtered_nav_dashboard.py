import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import altair as alt  # For more advanced charting


# Define the directory where the workbooks are stored
WORKBOOK_DIR = "NAV"  # Update this path to where your Excel workbooks are stored

# Function to list available Excel files in the specified directory
def list_workbooks(directory):
    try:
        # List only .xlsx files in the directory
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load NAV data from the selected workbook
def load_nav_data(file_path):
    try:
        # Read the first 10 columns (A-J) from the Excel file
        data = pd.read_excel(file_path, sheet_name=0, usecols="A:J")  # Load columns A-J
        
        # Check if 'Date' and 'NAV' columns exist for validation and charting purposes
        if 'NAV' not in data.columns or 'Date' not in data.columns:
            st.error("NAV or Date column not found in the selected workbook.")
            return pd.DataFrame()

        # Convert Date column to datetime format
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        data = data.sort_values(by='Date')  # Sort data by Date
        
        # Drop rows with missing Date or NAV
        data = data.dropna(subset=['Date', 'NAV'])

        return data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

# Function to filter data based on the selected date range
def filter_data_by_date(data, date_range):
    if date_range == "1 Day":
        return data.tail(1)
    elif date_range == "5 Days":
        return data.tail(5)
    elif date_range == "1 Month":
        one_month_ago = data['Date'].max() - timedelta(days=30)
        return data[data['Date'] >= one_month_ago]
    elif date_range == "6 Months":
        six_months_ago = data['Date'].max() - timedelta(days=180)
        return data[data['Date'] >= six_months_ago]
    elif date_range == "1 Year":
        one_year_ago = data['Date'].max() - timedelta(days=365)
        return data[data['Date'] >= one_year_ago]
    else:  # Max
        return data

# Function to recalculate NAV starting from 100
def recalculate_nav(filtered_data):
    # Start from an initial NAV value of 100
    initial_nav = filtered_data['NAV'].iloc[0]
    
    # Scale NAV values starting from 100
    filtered_data['Rebased NAV'] = (filtered_data['NAV'] / initial_nav) * 100
    return filtered_data

# Streamlit app layout and logic
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the directory
    workbooks = list_workbooks(WORKBOOK_DIR)

    # If no workbooks are found, display an error
    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    # Display dropdown menu to select a workbook
    selected_workbook = st.selectbox("Select a workbook", workbooks)

    # Date range options for the user
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    # Load and display NAV data from the selected workbook
    if selected_workbook:
        st.write(f"### Displaying data from {selected_workbook}")

        # Load NAV data (Columns A-J) from the selected workbook
        nav_data = load_nav_data(os.path.join(WORKBOOK_DIR, selected_workbook))
        
        # Check if NAV data is successfully loaded
        if not nav_data.empty:
            st.success("Data loaded successfully!")

            # Filter the data based on selected date range
            filtered_data = filter_data_by_date(nav_data, selected_range)

            # Remove the time from the Date column for cleaner display
            filtered_data['Date'] = filtered_data['Date'].dt.date

            # Recalculate NAV to start from 100 for ranges other than '1 Day' and '5 Days'
            if selected_range not in ["1 Day", "max"]:
                filtered_data = recalculate_nav(filtered_data)
                chart_column = 'Rebased NAV'
            else:
                chart_column = 'NAV'

            # Display the filtered data as a line chart using Altair, with y-axis starting from 80
            line_chart = alt.Chart(filtered_data).mark_line().encode(
                x='Date:T',
                y=alt.Y(f'{chart_column}:Q', scale=alt.Scale(domain=[80, filtered_data[chart_column].max()])),
                tooltip=['Date:T', f'{chart_column}:Q']
            ).properties(
                width=700,
                height=400
            )

            st.altair_chart(line_chart, use_container_width=True)

            # Rename column I as "Returns"
            if 'Unnamed: 8' in filtered_data.columns:
                filtered_data = filtered_data.rename(columns={'Unnamed: 8': 'Returns'})

            # Remove column B and rename column I as "Returns"
            filtered_data = filtered_data.drop(columns=['Stocks'], errors='ignore')  # Assuming 'Stocks' is in column B
            
            # Display the filtered data as a table (showing columns A-J, except B)
            st.write("### Data Table")
            st.dataframe(filtered_data.reset_index(drop=True))  # Reset index to remove the serial number

        else:
            st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
