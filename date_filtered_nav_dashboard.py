import streamlit as st
import pandas as pd
import os
from datetime import timedelta, datetime
import openpyxl

# Directory where the workbooks are stored
WORKBOOK_DIR = "NAV"  # Modify as per your setup

# Helper: List available workbooks
def list_workbooks(directory):
    try:
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Helper: Load the workbook into DataFrame
def load_workbook_data(file_path):
    """Loads the entire Excel workbook into a DataFrame."""
    try:
        data = pd.read_excel(file_path, sheet_name=0, dtype=str)  # Everything as string
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')  # Parse 'Date' column
        data = data.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)
        return data
    except Exception as e:
        st.error(f"Error loading workbook: {e}")
        return pd.DataFrame()

# Helper: Extract stock blocks based on "Stocks" and "Quantities" keywords
def extract_stock_blocks(data):
    """Extract all stock blocks and return as a list of (date, stock_names, block_data)."""
    stock_blocks = []
    current_block = None

    # Iterate through rows to identify 'Stocks' and 'Quantities' blocks
    for idx, row in data.iterrows():
        if row[1] == 'Stocks':  # Start of a new stock block
            stock_names = row[2:7].dropna().tolist()  # Extract stock names from C-G
            block_date = data.loc[idx - 1, 'Date']  # Date right before 'Stocks' row
            current_block = {
                "date": block_date,
                "stock_names": stock_names,
                "start_idx": idx
            }
        elif row[1] == 'Quantities' and current_block:  # End of the block
            # Capture the block data
            block_data = data.iloc[current_block["start_idx"]: idx + 1]
            stock_blocks.append((current_block["date"], current_block["stock_names"], block_data))
            current_block = None  # Reset for next block

    return stock_blocks

# Helper: Get relevant blocks based on selected time range
def get_relevant_blocks(stock_blocks, start_date, end_date):
    """Return blocks matching the date range or the latest block if none match."""
    relevant_blocks = [block for block in stock_blocks if start_date <= block[0] <= end_date]

    # If no blocks found, return the latest block (fallback)
    if not relevant_blocks and stock_blocks:
        relevant_blocks = [max(stock_blocks, key=lambda x: x[0])]  # Latest block by date

    return relevant_blocks

# Filter data based on date range
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

# Display relevant blocks in the dashboard
def display_relevant_blocks(relevant_blocks):
    """Displays relevant stock blocks in Streamlit."""
    for date, stock_names, block_data in relevant_blocks:
        st.write(f"### Stocks on {date.strftime('%Y-%m-%d')}")
        st.write(f"Stock Names: {', '.join(stock_names)}")
        st.dataframe(block_data.reset_index(drop=True))

# Main dashboard function
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the directory
    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No Excel workbooks found.")
        return

    # Select a workbook
    selected_workbook = st.selectbox("Select a Workbook", workbooks)
    file_path = os.path.join(WORKBOOK_DIR, selected_workbook)

    # Load the workbook data
    data = load_workbook_data(file_path)
    if data.empty:
        st.error("Failed to load data. Please check the workbook format.")
        return

    # Select a date range
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    # Filter the data by the selected range
    filtered_data = filter_data_by_date(data, selected_range)

    # Determine the start and end dates for stock extraction
    start_date = filtered_data['Date'].min()
    end_date = filtered_data['Date'].max()

    # Extract all stock blocks from the data
    stock_blocks = extract_stock_blocks(data)

    if not stock_blocks:
        st.warning("No stock blocks available in the workbook.")
        return

    # Get the relevant stock blocks for the selected time period
    relevant_blocks = get_relevant_blocks(stock_blocks, start_date, end_date)

    if not relevant_blocks:
        st.warning("No matching stock blocks found for the selected time range.")
    else:
        display_relevant_blocks(relevant_blocks)

if __name__ == "__main__":
    main()
