import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import altair as alt
import openpyxl
from datetime import datetime, timedelta
import yfinance as yf
import subprocess  # To run git commands
import time

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

        # Remove duplicate entries based on the 'Date' column
        data = data.drop_duplicates(subset='Date', keep='first')

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

# Function to recalculate NAV starting from 100
def recalculate_nav(filtered_data):
    # Start from an initial NAV value of 100
    initial_nav = filtered_data['NAV'].iloc[0]
    
    # Scale NAV values starting from 100
    filtered_data['Rebased NAV'] = (filtered_data['NAV'] / initial_nav) * 100
    return filtered_data

# Function to fetch stock data with retry logic
def fetch_stock_data_with_retry(stock_symbol, start, end, retries=3, delay=5):
    for i in range(retries):
        try:
            ticker = yf.Ticker(stock_symbol)
            hist = ticker.history(start=start, end=end, interval="1d", auto_adjust=False)
            if not hist.empty:
                return hist
        except Exception as e:
            print(f"Attempt {i+1} failed for {stock_symbol}. Error: {e}")
        time.sleep(delay)
    return pd.DataFrame()  # Return empty DataFrame if all retries fail

# Function to modify the Excel file locally
def modify_workbook(filename):
    # Path to the Excel file
    file_path = os.path.join(WORKBOOK_DIR, filename)
    
    try:
        # Load the Excel workbook using openpyxl
        workbook = openpyxl.load_workbook(file_path)

        # Modify all sheets in the workbook (you can insert your logic here)
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            print(f"Modifying sheet: {sheet_name}")

            # Step 1: Identify the last date in column A (assuming it's the date column)
            last_date_cell = ws.cell(row=ws.max_row, column=1).value
            if isinstance(last_date_cell, datetime):
                last_date = last_date_cell
            else:
                # If no valid date is found, set a default start date
                last_date = datetime.now() - timedelta(days=30)  # Assume 30 days ago if no valid date
            next_date = last_date + timedelta(days=1)  # Next date after the last date in the sheet

            # Step 2: Identify the last non-zero NAV in column J (NAV)
            nav_column_index = 10  # Column J for NAV
            last_non_zero_nav = None
            last_nav_row = None

            for row in range(ws.max_row, 2, -1):
                nav_value = ws.cell(row=row, column=nav_column_index).value
                if isinstance(nav_value, (int, float)) and nav_value != 0:
                    last_non_zero_nav = nav_value
                    last_nav_row = row
                    break

            if last_non_zero_nav is None:
                last_non_zero_nav = 100  # Default NAV value
                last_nav_row = 3

            # Step 3: Identify existing stock symbols and quantities in columns C to G
            stocks_row = None
            quantities_row = None

            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value == "Stocks":
                    stocks_row = row
                elif cell_value == "Quantities":
                    quantities_row = row

            if not stocks_row or not quantities_row:
                print(f"Could not find 'Stocks' or 'Quantities' headers in sheet {sheet_name}. Skipping sheet.")
                continue

            stocks = {}
            quantities = []

            for col in range(3, 8):  # Columns C to G
                stock_symbol = ws.cell(row=stocks_row, column=col).value
                quantity = ws.cell(row=quantities_row, column=col).value
                if stock_symbol and isinstance(stock_symbol, str):
                    stocks[stock_symbol] = stock_symbol  # Use stock symbol as stock name
                    quantities.append(quantity)

            # Step 4: Fetch historical stock data from the next date after the last date in the sheet until today
            today_date = datetime.now().strftime('%Y-%m-%d')
            next_date_str = next_date.strftime('%Y-%m-%d')

            all_prices = {}
            closing_dates = []

            for stock_symbol in stocks.keys():
                hist = fetch_stock_data_with_retry(stock_symbol, next_date_str, today_date)
                if not hist.empty:
                    closing_prices = hist['Close'].tolist()
                    closing_dates = hist.index.strftime('%Y-%m-%d').tolist()
                    all_prices[stock_symbol] = (closing_dates, closing_prices)

            # Step 5: Insert the fetched data and perform calculations
            current_row = ws.max_row + 1

            basket_values = []
            returns = []
            nav_values = [last_non_zero_nav]

            for i in range(len(closing_dates)):
                ws.cell(row=current_row + i, column=1, value=closing_dates[i])  # Insert date

                basket_value = 0
                for j, stock_symbol in enumerate(stocks.keys()):
                    price = all_prices[stock_symbol][1][i] if i < len(all_prices[stock_symbol][1]) else 0
                    quantity = quantities[j]
                    basket_value += price * quantity
                    ws.cell(row=current_row + i, column=3 + j, value=price)  # Insert price starting from column C

                ws.cell(row=current_row + i, column=8, value=basket_value)  # Insert basket value

                basket_values.append(basket_value)

                ret = (basket_value - basket_values[i - 1]) / basket_values[i - 1] if i > 0 and basket_values[i - 1] != 0 else 0
                returns.append(ret)
                ws.cell(row=current_row + i, column=9, value=ret)  # Insert return

                nav = nav_values[-1] * (1 + ret)
                nav_values.append(nav)
                ws.cell(row=current_row + i, column=10, value=nav)  # Insert NAV

        # Save the modified Excel file locally
        workbook.save(file_path)
        print(f"Workbook {filename} has been successfully modified and saved.")

    except Exception as e:
        st.error(f"Error modifying {filename}: {e}")

# Function to execute git commands to add, commit, and push changes
def git_add_commit_push(workbooks):
    try:
        # Add each modified workbook individually to ensure all changes are tracked
        for workbook in workbooks:
            subprocess.run(["git", "add", os.path.join(WORKBOOK_DIR, workbook)], check=True)

        # Git commit with a message
        commit_message = f"Updated all workbooks with new data"
        subprocess.run(["git", "commit", "-m", commit_message], check=True)

        # Git push to the remote repository
        subprocess.run(["git", "push"], check=True)
        print("All changes have been successfully pushed to GitHub.")

    except subprocess.CalledProcessError as e:
        print(f"Error during git operation: {e}")

# Function to modify and push changes to all Excel files in the directory
def modify_all_workbooks():
    workbooks = list_workbooks(WORKBOOK_DIR)
    for workbook in workbooks:
        modify_workbook(workbook)

    # After modifying all workbooks, push changes to GitHub
    git_add_commit_push(workbooks)

# Streamlit app layout and logic
def main():
    st.title("NAV Data Dashboard")

    # Modify and push changes to all Excel files
    if st.button("Update All Workbooks"):
        modify_all_workbooks()

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

    if selected_workbook:
        st.write(f"### Displaying data from {selected_workbook}")

        # Load NAV data (Columns A-J) from the selected workbook
        nav_data = load_nav_data(os.path.join(WORKBOOK_DIR, selected_workbook))
        
        # Check if NAV data is successfully loaded
        if not nav_data.empty:
            # Filter the data based on selected date range
            filtered_data = filter_data_by_date(nav_data, selected_range)

            # Remove the time from the Date column for cleaner display
            filtered_data['Date'] = filtered_data['Date'].dt.date

            # Recalculate NAV to start from 100 for ranges other than '1 Day' and '5 Days'
            if selected_range not in ["1 Day", "Max"]:
                filtered_data = recalculate_nav(filtered_data)
                chart_column = 'Rebased NAV'
            else:
                chart_column = 'NAV'

            # Display the filtered data as a line chart using Altair, with y-axis starting from 80
            line_chart = alt.Chart(filtered_data).mark_line().encode(
                x='Date:T',
                y=alt.Y(f'{chart_column}:Q', scale=alt.Scale(domain=[80, filtered_data[chart_column].max()])),
                tooltip=['Date:T', f'{chart_column}:Q']
            ).properties(
                width=700,
                height=400
            )

            st.altair_chart(line_chart, use_container_width=True)

            # Display the filtered data as a table (showing columns A-J, except B)
            st.write("### Data Table")
            st.dataframe(filtered_data.reset_index(drop=True))  #
            # Reset index to remove the serial number
            st.dataframe(filtered_data.reset_index(drop=True)) 

        else:
            st.error("Failed to load data. Please check the workbook format.")

if __name__ == "__main__":
    main()
