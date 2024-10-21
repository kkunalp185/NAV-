import streamlit as st
import pandas as pd
import os
from datetime import timedelta, datetime
import openpyxl

# Define the directory where the workbooks are stored (this is in the same repo)
WORKBOOK_DIR = "NAV"  # Folder where the Excel workbooks are stored

# Function to list available Excel files in the specified directory
def list_workbooks(directory):
    try:
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load NAV data from the selected workbook and handle date parsing
def load_nav_data(file_path):
    try:
        # Load full sheet data without limiting columns and headers
        data = pd.read_excel(file_path, sheet_name=0, header=None)
        data.columns = ['Date', 'Header', 'Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5', 'Basket Value', 'Returns', 'NAV']
        
        # Drop the 'Header' column as per the request
        data.drop('Header', axis=1, inplace=True)
        
        # Ensure 'Date' column is datetime; coerce errors to handle non-date values
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        
        return data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

# Function to filter data based on the selected date range
def filter_data_by_date(data, date_range):
    if 'Date' not in data.columns:
        st.error("Date column not found in the data for filtering.")
        return data

    data = data.dropna(subset=['Date'])

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
    else:
        return data

# Function to process Excel data and identify stock name changes dynamically
def process_excel_data(data):
    stock_blocks = []
    current_block = None

    # Iterate through the rows of the DataFrame to detect stock changes
    for idx, row in data.iterrows():
        if isinstance(row['Stock1'], str) and row['Stock1'] == 'Stocks':  # Detect when stock names change
            if current_block:
                current_block['end_idx'] = idx - 1  # End the current block before the next 'Stocks' row
                stock_blocks.append(current_block)  # Save the completed block

            # Create a new block
            stock_names = row[['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']].tolist()  # Get stock names from columns
            current_block = {'stock_names': stock_names, 'start_idx': idx + 2, 'end_idx': None, 'dates': []}  # Include dates
            st.write(f"DEBUG: Fetched stock names: {stock_names}")

    if current_block:
        current_block['end_idx'] = len(data) - 1  # Handle the last block until the end of the dataset
        stock_blocks.append(current_block)

    # Add dates to each stock block
    for block in stock_blocks:
        block['dates'] = data.iloc[block['start_idx']:block['end_idx'] + 1].dropna(subset=['Date'])['Date'].tolist()
        st.write(f"DEBUG: Dates for block: {block['dates']}")

    return stock_blocks

# Function to insert stock names for the relevant block above the selected time period's data
def insert_stock_names_above_data(stock_blocks, filtered_data):
    final_data = pd.DataFrame()

    filtered_dates = filtered_data['Date'].tolist()

    for block in stock_blocks:
        # Check if any dates from this block overlap with filtered dates
        overlap_dates = [date for date in block['dates'] if date in filtered_dates]

        # If there are overlapping dates, insert the stock names above the first relevant date
        if overlap_dates:
            # Insert the stock names once, before the first matching date in the block
            stock_names_row = pd.DataFrame([[None] * len(filtered_data.columns)], columns=filtered_data.columns)
            for i, stock_name in enumerate(block['stock_names']):
                stock_names_row[f'Stock{i + 1}'] = stock_name

            # Add the stock names row to the final data
            final_data = pd.concat([final_data, stock_names_row], ignore_index=True)

            # Add the block's data that overlaps with the filtered data, without duplicate date entries
            block_data = filtered_data[filtered_data['Date'].isin(overlap_dates)]
            final_data = pd.concat([final_data, block_data.drop_duplicates(subset=['Date'], keep='first')], ignore_index=True)

    return final_data

# Main Streamlit app function
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the directory
    workbooks = list_workbooks(WORKBOOK_DIR)

    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    # Display the data for a specific workbook (example: the first one)
    selected_workbook = st.selectbox("Select a workbook", workbooks)
    
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    nav_data = load_nav_data(file_path)

    if not nav_data.empty:
        # Process the Excel data and detect stock name changes, including block dates
        stock_blocks = process_excel_data(nav_data)

        if not stock_blocks:
            st.error("No valid stock data found in the workbook.")
            return

        # Allow the user to select a date range
        date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
        selected_range = st.selectbox("Select Date Range", date_ranges)

        # Filter the nav_data by the selected date range
        filtered_data = filter_data_by_date(nav_data, selected_range)

        # Insert stock names above the relevant block data
        final_data = insert_stock_names_above_data(stock_blocks, filtered_data)

        # Display the combined filtered data in a single table
        st.write("### Combined Stock Data Table")
        st.dataframe(final_data.reset_index(drop=True))

    else:
        st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
