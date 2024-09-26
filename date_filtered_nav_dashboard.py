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

            # Ensure that we only use Date and NAV columns for the chart
            filtered_chart_data = filtered_data[['Date', 'NAV']]

            # Display the filtered data as a line chart using Altair, with y-axis starting from 80
            line_chart = alt.Chart(filtered_chart_data).mark_line().encode(
                x='Date:T',
                y=alt.Y('NAV:Q', scale=alt.Scale(domain=[80, filtered_chart_data['NAV'].max()])),
                tooltip=['Date:T', 'NAV:Q']
            ).properties(
                width=700,
                height=400
            )

            st.altair_chart(line_chart, use_container_width=True)

            # Remove column B and rename column 8 as "Returns" (if column 8 exists)
            if ' 8' in filtered_data.columns:  # Replace with the exact column name if it differs
                filtered_data = filtered_data.rename(columns={' 8': 'Returns'})
            filtered_data = filtered_data.drop(columns=['Stocks'], errors='ignore')  # Assuming 'Stocks' is in column B
            
            # Display the filtered data as a table (showing columns A-J, except B)
            st.write("### Data Table, without 'Stocks')")
            st.dataframe(filtered_data.reset_index(drop=True))  # Reset index to remove the serial number

        else:
            st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
