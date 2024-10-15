import streamlit as st
import pandas as pd
import os
from datetime import timedelta, datetime
import openpyxl

# Directory where the workbooks are stored
WORKBOOK_DIR = "NAV"

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
    try:
        data = pd.read_excel(file_path, sheet_name=0, dtype=str)  # Load as string for safety
        data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Trim spaces
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        data = data.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)
        return data
    except Exception as e:
        st.error(f"Error loading workbook: {e}")
        return pd.DataFrame()

# Helper: Extract stock blocks
def extract_stock_blocks(data):
    stock_blocks = []
    current_block = None

    for idx, row in data.iterrows():
        # Detect 'Stocks' (case-insensitive and trimmed)
        if str(row[1]).strip().lower() == 'stocks':
            stock_names = row[2:7].dropna().tolist()  # Extract stock names from columns C-G
            block_date = data.loc[idx - 1, 'Date']  # Date from the row above
            st.write(f"Found stock block at index {idx} with date {block_date}.")  # Debug output
            current_block = {
                "date": block_date,
                "stock_names": stock_names,
                "start_idx": idx
            }

        # Detect 'Quantities' row
        elif str(row[1]).strip().lower() == 'quantities' and current_block:
            block_data = data.iloc[current_block["start_idx"]: idx + 1]  # Extract the block
            stock_blocks.append((current_block["date"], current_block["stock_names"], block_data))
            current_block = None  # Reset for the next block

    st.write(f"Total Stock Blocks Found: {len(stock_blocks)}")  # Debug output
    return stock_blocks

# Helper: Get relevant blocks based on selected time range
def get_relevant_blocks(stock_blocks, start_date, end_date):
    relevant_blocks = [block for block in stock_blocks if start_date <= block[0] <= end_date]

    if not relevant_blocks and stock_blocks:
        relevant_blocks = [max(stock_blocks, key=lambda x: x[0])]  # Latest block

    return relevant_blocks

# Filter data by date range
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
    else:
        return data

# Display relevant blocks
def display_relevant_blocks(relevant_blocks):
    for date, stock_names, block_data in relevant_blocks:
        st.write(f"### Stocks on {date.strftime('%Y-%m-%d')}")
        st.write(f"Stock Names: {', '.join(stock_names)}")
        st.dataframe(block_data.reset_index(drop=True))

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

    # Load workbook data
    data = load_workbook_data(file_path)
    if data.empty:
        st.error("Failed to load data. Please check the workbook format.")
        return

    # Select date range
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    # Filter data by date range
    filtered_data = filter_data_by_date(data, selected_range)
    start_date, end_date = filtered_data['Date'].min(), filtered_data['Date'].max()

    # Extract stock blocks
    stock_blocks = extract_stock_blocks(data)

    if not stock_blocks:
        st.warning("No stock blocks available in the workbook.")
        return

    # Get relevant stock blocks
    relevant_blocks = get_relevant_blocks(stock_blocks, start_date, end_date)

    if not relevant_blocks:
        st.warning("No matching stock blocks found for the selected time range.")
    else:
        display_relevant_blocks(relevant_blocks)

if __name__ == "__main__":
    main()
