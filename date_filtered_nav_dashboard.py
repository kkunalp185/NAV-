import streamlit as st
import pandas as pd
import os
from datetime import timedelta
from datetime import datetime
import openpyxl

# Define the directory where the workbooks are stored (this is in the same repo)
WORKBOOK_DIR = "NAV"  # Folder where the Excel workbooks are stored

# Function to list available Excel files in the specified directory
def list_workbooks(directory):
    try:
        # List only .xlsx files in the directory
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load NAV data from the selected workbook and handle date parsing
def load_nav_data(file_path):
    try:
        # Load the data and skip the first row if it contains extra headers
        data = pd.read_excel(file_path, sheet_name=0, header=1)  # Skip the first row if it contains headers
        
        # Ensure 'Date' column is datetime; coerce errors to handle non-date values
        if 'Date' in data.columns:
            data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
            # Drop rows where 'Date' couldn't be parsed (NaT values)
            data = data.dropna(subset=['Date'])
        else:
            st.error("Date column not found in the dataset.")
        return data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

# Function to filter data based on the selected date range
def filter_data_by_date(data, date_range):
    if 'Date' not in data.columns:
        st.error("Date column not found in the data for filtering.")
        return data

    # Ensure all 'Date' values are valid datetime objects
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
    else:  # Max
        return data

# Function to process the Excel data and identify stock name changes dynamically
def process_excel_data(data):
    stock_blocks = []
    current_block = None

    # Dynamically find the column that contains 'Stocks'
    stock_column = None
    for col in data.columns:
        if data[col].astype(str).str.contains('Stocks').any():
            stock_column = col
            break

    if not stock_column:
        st.error("No 'Stocks' column found in the workbook.")
        return []

    # Iterate through the rows of the DataFrame
    for idx, row in data.iterrows():
        if isinstance(row[stock_column], str) and row[stock_column] == 'Stocks':  # Detect when stock names change
            if current_block:
                current_block['end_idx'] = idx - 1  # End the current block before the next 'Stocks' row
                stock_blocks.append(current_block)  # Save the completed block

            # Create a new block
            stock_names = row[2:7].tolist()  # Get stock names from columns C to G
            current_block = {'stock_names': stock_names, 'start_idx': idx + 2, 'end_idx': None}

    if current_block:
        current_block['end_idx'] = len(data) - 1  # Handle the last block until the end of the dataset
        stock_blocks.append(current_block)

    # Create a combined DataFrame to store all the blocks
    combined_data = pd.DataFrame()

    # Rename stock columns to Stock1, Stock2, etc. and process blocks of data
    for block in stock_blocks:
        block_data = data.iloc[block['start_idx']:block['end_idx'] + 1].copy()
        stock_columns = ['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']

        # Map original stock names to Stock1, Stock2, etc.
        column_mapping = {data.columns[i]: stock_columns[i - 2] for i in range(2, 7)}

        # Rename columns in the block data
        block_data = block_data.rename(columns=column_mapping)

        # Insert a row with the stock names before the data for the block
        stock_names_row = pd.DataFrame({
            'Stock1': [block['stock_names'][0]],
            'Stock2': [block['stock_names'][1]],
            'Stock3': [block['stock_names'][2]],
            'Stock4': [block['stock_names'][3]],
            'Stock5': [block['stock_names'][4]],
            'Date': [None], 'Basket Value': [None], 'Returns': [None], 'NAV': [None]
        })

        # Concatenate stock names row with the block data
        block_data = pd.concat([stock_names_row, block_data], ignore_index=True)

        # Append to the combined DataFrame
        combined_data = pd.concat([combined_data, block_data], ignore_index=True)

    return combined_data

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

        # Display the combined filtered data in a single table
        st.write("### Combined Stock Data Table")
        st.dataframe(filtered_data.reset_index(drop=True))

    else:
        st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
