import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import openpyxl
from datetime import datetime

# Define the directory where the workbooks are stored
WORKBOOK_DIR = "NAV"

# Function to list available Excel files in the specified directory
def list_workbooks(directory):
    try:
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load the entire worksheet data
def load_entire_workbook(file_path):
    """Load the entire worksheet as a DataFrame without any modifications."""
    try:
        data = pd.read_excel(file_path, sheet_name=0, dtype=str)  # Load everything as string to preserve formatting
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')  # Parse 'Date' column if it exists
        data = data.dropna(subset=['Date'])  # Drop rows without a valid date
        data = data.sort_values('Date').reset_index(drop=True)  # Sort by 'Date'
        return data
    except Exception as e:
        st.error(f"Error loading workbook: {e}")
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

# Main function
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks
    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No Excel workbooks found.")
        return

    # Select a workbook
    selected_workbook = st.selectbox("Select a Workbook", workbooks)
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    # Load the entire worksheet data
    data = load_entire_workbook(file_path)
    if data.empty:
        st.error("Failed to load data. Please check the workbook format.")
        return

    # Select a date range
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    # Filter the data by the selected date range
    filtered_data = filter_data_by_date(data, selected_range)

    # Display the filtered data in a table
    st.write("### Data Table")
    st.dataframe(filtered_data)

if __name__ == "__main__":
    main()
