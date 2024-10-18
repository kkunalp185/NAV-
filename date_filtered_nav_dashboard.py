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
        data = pd.read_excel(file_path, sheet_name=0)  # Load the full sheet data
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
    data['Date'] = pd.to_datetime(data['Date'], errors='coerce')

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

# Function to process the Excel data and identify stock blocks dynamically
def process_excel_data(data):
    stock_blocks = []
    current_block = None

    # Iterate through the rows to detect stock and quantity rows
    for idx, row in data.iterrows():
        if isinstance(row['Stocks'], str) and row['Stocks'] == 'Stocks':  # Detect the stock row
            if current_block:
                current_block['end_idx'] = idx - 1  # End the current block before the next 'Stocks' row
                stock_blocks.append(current_block)  # Save the completed block

            # Create a new block for the current stock configuration
            stock_names = row[2:7].tolist()  # Get stock names from columns C to G
            quantities_row = data.iloc[idx + 1]  # The next row should be the quantities row
            quantities = quantities_row[2:7].tolist()  # Get quantities from columns C to G

            current_block = {
                'stock_names': stock_names,
                'quantities': quantities,
                'start_idx': idx + 2,  # Data starts from two rows after the stock names
                'end_idx': None
            }

    if current_block:
        current_block['end_idx'] = len(data) - 1  # Handle the last block until the end of the dataset
        stock_blocks.append(current_block)

    # Create a combined DataFrame to store all the blocks
    combined_data = pd.DataFrame()

    # Process and rename columns for each stock block
    for block in stock_blocks:
        block_data = data.iloc[block['start_idx']:block['end_idx'] + 1].copy()

        # Rename stock columns to Stock1, Stock2, etc.
        stock_columns = ['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']
        column_mapping = {data.columns[i]: stock_columns[i - 2] for i in range(2, 7)}
        block_data = block_data.rename(columns=column_mapping)

        # Create a row with stock names and quantities
        stock_names_row = pd.DataFrame({
            'Stock1': [block['stock_names'][0]],
            'Stock2': [block['stock_names'][1]],
            'Stock3': [block['stock_names'][2]],
            'Stock4': [block['stock_names'][3]],
            'Stock5': [block['stock_names'][4]],
            'Date': [None], 'Basket Value': [None], 'Returns': [None], 'NAV': [None]
        })
        quantities_row = pd.DataFrame({
            'Stock1': [f"Qty: {block['quantities'][0]}"],
            'Stock2': [f"Qty: {block['quantities'][1]}"],
            'Stock3': [f"Qty: {block['quantities'][2]}"],
            'Stock4': [f"Qty: {block['quantities'][3]}"],
            'Stock5': [f"Qty: {block['quantities'][4]}"],
            'Date': [None], 'Basket Value': [None], 'Returns': [None], 'NAV': [None]
        })

        # Concatenate the stock names and quantities row with the block data
        block_data = pd.concat([stock_names_row, quantities_row, block_data], ignore_index=True)

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

        # Ensure stock names for the latest block are displayed at the start
        if not filtered_data.empty:
            st.write("### Combined Stock Data Table")
            st.dataframe(filtered_data.reset_index(drop=True))
        else:
            st.error("No data found for the selected date range.")

    else:
        st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
