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
        data = pd.read_excel(file_path, sheet_name=0, header=None)
        data.columns = data.iloc[0]  # Use the first row as headers
        data = data.drop(0).reset_index(drop=True)  # Drop the header row and reset index
        if 'Date' in data.columns:
            data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
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

    # Find the column that contains 'Stocks'
    stock_column = None
    for col in data.columns:
        if data[col].astype(str).str.contains('Stocks').any():
            stock_column = col
            break

    if not stock_column:
        st.error("No 'Stocks' column found in the workbook.")
        return []

    # Identify stock name blocks
    for idx, row in data.iterrows():
        if isinstance(row[stock_column], str) and row[stock_column] == 'Stocks':
            if current_block:
                current_block['end_idx'] = idx - 1
                stock_blocks.append(current_block)

            stock_names = row[2:7].tolist()  # Get stock names from columns C to G
            current_block = {'stock_names': stock_names, 'start_idx': idx + 2, 'end_idx': None}
            st.write(f"DEBUG: Fetched stock names: {stock_names}")

    if current_block:
        current_block['end_idx'] = len(data) - 1
        stock_blocks.append(current_block)

    # Create combined DataFrame for all blocks
    combined_data = pd.DataFrame()

    for block in stock_blocks:
        block_data = data.iloc[block['start_idx']:block['end_idx'] + 1].copy()
        block_dates = block_data['Date'].tolist()
        st.write(f"DEBUG: Dates for block: {block_dates}")

        stock_columns = ['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']

        # Rename stock columns in the block
        column_mapping = {data.columns[i]: stock_columns[i - 2] for i in range(2, 7)}
        block_data = block_data.rename(columns=column_mapping)

        # Create stock names row
        stock_names_row = pd.DataFrame([[None] * len(block_data.columns)], columns=block_data.columns)
        for i, stock_name in enumerate(block['stock_names']):
            stock_names_row[f'Stock{i + 1}'] = stock_name

        # Insert stock names row just before block data
        combined_data = pd.concat([combined_data, stock_names_row, block_data], ignore_index=True)

    return combined_data

# Function to insert stock names for the relevant block above the selected time period's data
def insert_stock_names_above_data(combined_data, filtered_data):
    final_data = pd.DataFrame()
    last_inserted_block = None

    # Get the first date in the filtered data to determine the relevant block
    first_filtered_date = filtered_data['Date'].min()

    for idx, row in combined_data.iterrows():
        # If the row is a stock names row (without date)
        if pd.isna(row['Date']):
            current_block = row[['Stock1', 'Stock2', 'Stock3', 'Stock4', 'Stock5']].values.tolist()

            # Check if the current block contains the first date from filtered data
            block_data = combined_data.loc[idx + 1:]  # Get data of the current block
            block_data_dates = block_data.dropna(subset=['Date'])['Date'].tolist()

            # Insert stock names only if the first_filtered_date is in the current block
            if first_filtered_date in block_data_dates:
                if last_inserted_block != current_block:
                    final_data = pd.concat([final_data, row.to_frame().T], ignore_index=True)
                    last_inserted_block = current_block

        # Append data rows to the final data if dates match the filtered dates
        if row['Date'] in filtered_data['Date'].tolist():
            final_data = pd.concat([final_data, row.to_frame().T], ignore_index=True)

    return final_data

# Main Streamlit app function
def main():
    st.title("NAV Data Dashboard")

    workbooks = list_workbooks(WORKBOOK_DIR)

    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    selected_workbook = st.selectbox("Select a workbook", workbooks)
    
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    nav_data = load_nav_data(file_path)

    if not nav_data.empty:
        combined_data = process_excel_data(nav_data)

        if combined_data.empty:
            st.error("No valid stock data found in the workbook.")
            return

        date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
        selected_range = st.selectbox("Select Date Range", date_ranges)

        filtered_data = filter_data_by_date(combined_data, selected_range)

        # Insert stock names above the relevant block data
        final_data = insert_stock_names_above_data(combined_data, filtered_data)

        # Display the final data in a single table
        st.write("### Combined Stock Data Table")
        st.dataframe(final_data.reset_index(drop=True))

    else:
        st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
