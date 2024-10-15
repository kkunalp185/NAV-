import streamlit as st
import pandas as pd
import os
from datetime import timedelta, datetime
import openpyxl

# Define the directory where the workbooks are stored
WORKBOOK_DIR = "NAV"  # Folder where the Excel workbooks are stored

def list_workbooks(directory):
    try:
        # List only .xlsx files in the directory
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

def load_workbook_data(file_path):
    """Loads the entire sheet into a DataFrame."""
    try:
        data = pd.read_excel(file_path, sheet_name=0, dtype=str)
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        data = data.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)
        return data
    except Exception as e:
        st.error(f"Error loading workbook: {e}")
        return pd.DataFrame()

def extract_stock_blocks(data):
    """
    Extracts all 'Stocks' and 'Quantities' blocks from the data,
    returning a list of tuples (block_date, stock_names, block_data).
    """
    stock_blocks = []
    current_block = None

    for idx, row in data.iterrows():
        if row[1] == 'Stocks':
            stock_names = row[2:7].dropna().tolist()  # Extract stock names from columns C to G
            block_date = data.loc[idx - 1, 'Date']  # Date just above the 'Stocks' row
            current_block = {
                "date": block_date,
                "stock_names": stock_names,
                "start_idx": idx
            }
        elif row[1] == 'Quantities' and current_block:
            block_data = data.iloc[current_block["start_idx"]: idx + 1]  # Capture the block data
            stock_blocks.append((current_block["date"], current_block["stock_names"], block_data))
            current_block = None  # Reset for the next block

    return stock_blocks

def get_relevant_blocks(stock_blocks, start_date, end_date):
    """
    Filters stock blocks based on the given date range.
    If no block matches, return the latest block.
    """
    relevant_blocks = [block for block in stock_blocks if start_date <= block[0] <= end_date]

    if not relevant_blocks:  # If no relevant blocks found, return the latest block
        relevant_blocks = [max(stock_blocks, key=lambda x: x[0])]

    return relevant_blocks

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

def display_relevant_blocks(relevant_blocks):
    """Displays the relevant stock blocks in Streamlit."""
    for date, stock_names, block_data in relevant_blocks:
        st.write(f"### Stocks on {date.strftime('%Y-%m-%d')}")
        st.write(f"Stock Names: {', '.join(stock_names)}")
        st.dataframe(block_data.reset_index(drop=True))

def main():
    st.title("NAV Data Dashboard")

    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No workbooks found.")
        return

    selected_workbook = st.selectbox("Select a Workbook", workbooks)
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    data = load_workbook_data(file_path)
    if data.empty:
        st.error("Failed to load data. Please check the workbook format.")
        return

    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    filtered_data = filter_data_by_date(data, selected_range)
    start_date = filtered_data['Date'].min()
    end_date = filtered_data['Date'].max()

    stock_blocks = extract_stock_blocks(data)
    relevant_blocks = get_relevant_blocks(stock_blocks, start_date, end_date)

    display_relevant_blocks(relevant_blocks)

if __name__ == "__main__":
    main()
