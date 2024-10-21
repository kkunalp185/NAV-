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
        if isinstance(row['Header'], str) and row['Header'] == 'Stocks':  # Detect when stock names change
            if current_block:
                current_block['end_idx'] = idx - 1  # End the current block before the next 'Stocks' row
                stock_blocks.append(current_block)  # Save the completed block

            # Create a new block
            stock_names = row[['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']].tolist()  # Get stock names from columns C to G
            current_block = {'stock_names': stock_names, 'start_idx': idx + 2, 'end_idx': None}  # Skip 'Quantities' row
            st.write(f"DEBUG: Fetched stock names: {stock_names}")

    if current_block:
        current_block['end_idx'] = len(data) - 1  # Handle the last block until the end of the dataset
        stock_blocks.append(current_block)

    # Create a combined DataFrame to store all the blocks
    combined_data = pd.DataFrame()

    # Rename stock columns to Stock1, Stock2, etc. and process blocks of data
    for block in stock_blocks:
        block_data = data.iloc[block['start_idx']:block['end_idx'] + 1].copy()

        # Skip the 'Quantities' row
        block_data = block_data[block_data['Header'] != 'Quantities']

        # Insert stock names once before the block's data
        stock_names_row = pd.DataFrame([[None] * len(block_data.columns)], columns=block_data.columns)
        for i, stock_name in enumerate(block['stock_names']):
            stock_names_row[f'Stock{i + 1}'] = stock_name

        # Add the stock names row and the block data to the combined DataFrame
        combined_data = pd.concat([combined_data, stock_names_row, block_data], ignore_index=True)

    return combined_data

# Function to insert stock names for the relevant block above the selected time period's data
def insert_stock_names_above_data(combined_data, filtered_data):
    final_data = pd.DataFrame()
    last_inserted_block = None

    # Find stock blocks that match the filtered data's dates
    filtered_dates = filtered_data['Date'].tolist()

    for idx, row in combined_data.iterrows():
        # If the row is a stock names row (without date)
        if pd.isna(row['Date']):
            current_block = row[['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']].values.tolist()

            # Insert stock names only once per block, when the first relevant date in the block appears
            if last_inserted_block != current_block:
                block_data = combined_data.loc[idx + 1:]  # Get data of the current block
                block_data_dates = block_data.dropna(subset=['Date'])['Date'].tolist()

                # Check if any of the block's dates overlap with filtered dates
                overlap_dates = [d for d in block_data_dates if d in filtered_dates]

                if overlap_dates:
                    # Insert stock names row just above the first overlap date
                    final_data = pd.concat([final_data, row.to_frame().T], ignore_index=True)
                    last_inserted_block = current_block

        # Append data rows to the final data, as long as dates match the filtered dates
        if row['Date'] in filtered_dates:
            final_data = pd.concat([final_data, row.to_frame().T], ignore_index=True)

    return final_data

# Main Streamlit app function
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the directory
    workbooks = list_workbooks(WORKBOOK_DIR)

    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    # Display the data for a specific workbook
    selected_workbook = st.selectbox("Select a workbook", workbooks)
    
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    nav_data = load_nav_data(file_path)

    if not nav_data.empty:
        # Process the Excel data and detect stock name changes (combine into a single table)
        combined_data = process_excel_data(nav_data)

        if combined_data.empty:
            st.error("No valid stock data found in the workbook.")
            return

        # Allow the user to select a date range
        date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
        selected_range = st.selectbox("Select Date Range", date_ranges)

        # Filter the combined data by the selected date range
        filtered_data = filter_data_by_date(combined_data, selected_range)

        # Check that the filtered data is not empty
        if filtered_data.empty:
            st.warning("No data available for the selected time period.")
            return

        # Create a new dataframe to insert stock names only once per block
        final_data = pd.DataFrame()
        current_block_start = None

        for block in stock_blocks:
            # Filter block dates to find relevant data for the selected range
            block_data = filter_data_by_date(block['data'], selected_range)

            # If there is no data in this block for the selected date range, skip it
            if block_data.empty:
                continue

            # Insert stock names only once above the first row of the block
            if current_block_start is None or block_data['Date'].iloc[0] != current_block_start:
                # Insert a row with stock names under the appropriate columns
                stock_names_row = pd.DataFrame({
                    'Stock1': [block['stock_names'][0]],
                    'Stock2': [block['stock_names'][1]],
                    'Stock3': [block['stock_names'][2]],
                    'Stock4': [block['stock_names'][3]],
                    'Stock5': [block['stock_names'][4]],
                    'Date': [None],
                    'Basket Value': [None],
                    'Returns': [None],
                    'NAV': [None]
                })
                final_data = pd.concat([final_data, stock_names_row], ignore_index=True)

                # Set the start date of the block to avoid repeated stock name insertion
                current_block_start = block_data['Date'].iloc[0]

            # Append the block data after inserting stock names
            final_data = pd.concat([final_data, block_data], ignore_index=True)

        # Display the final combined data
        st.write("### Combined Stock Data Table")
        st.dataframe(final_data.reset_index(drop=True))

    else:
        st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
