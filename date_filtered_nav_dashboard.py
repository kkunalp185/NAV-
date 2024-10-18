import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import altair as alt
import openpyxl
from openpyxl.styles import NamedStyle
from datetime import datetime
import yfinance as yf
import subprocess
from openpyxl.utils import get_column_letter # To run git commands

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

def load_full_data(file_path):
    try:
        # Load the entire workbook (without limiting columns to A:J)
        data = pd.read_excel(file_path, sheet_name=0)  # Load the entire sheet without filtering columns
        # Ensure the 'Date' column is properly converted to datetime
        if 'Date' in data.columns:
            data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        return data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def load_nav_data(file_path):
    try:
        data = pd.read_excel(file_path, sheet_name=0)  # Load full sheet data without limiting columns
        # Ensure 'Date' column is datetime; coerce errors to handle non-date values
        if 'Date' in data.columns:
            data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        else:
            st.error("Date column not found in the dataset.")
        return data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

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

# Function to load NAV data from the selected workbook
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

    # Process blocks of data and insert stock names
    for block in stock_blocks:
        block_data = data.iloc[block['start_idx']:block['end_idx'] + 1].copy()

        # Create a row for stock names and leave the rest as None
        stock_row = pd.DataFrame({
            'Date': [None], 
            'Stock1': [block['stock_names'][0]], 
            'Stock2': [block['stock_names'][1]], 
            'Stock3': [block['stock_names'][2]], 
            'Stock4': [block['stock_names'][3]], 
            'Stock5': [block['stock_names'][4]], 
            'Basket Value': [None], 
            'Returns': [None], 
            'NAV': [None]
        })

        # Append stock name row followed by actual data
        block_data = pd.concat([stock_row, block_data], ignore_index=True)

        # Append to the combined DataFrame
        combined_data = pd.concat([combined_data, block_data], ignore_index=True)

    return combined_data

# Function to recalculate NAV starting from 100
def recalculate_nav(filtered_data):
    initial_nav = filtered_data['NAV'].iloc[0]
    filtered_data['Rebased NAV'] = (filtered_data['NAV'] / initial_nav) * 100
    return filtered_data

# Function to modify all Excel files in the directory and push them to GitHub
def modify_all_workbooks_and_push_to_github():
    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No workbooks found to modify.")
        return

    modified_files = []

    for filename in workbooks:
        try:
            modify_workbook(filename)
            modified_files.append(filename)
        except Exception as e:
            st.error(f"Error modifying {filename}: {e}")

    # Push all modified files to GitHub
    if modified_files:
        git_add_commit_push(modified_files)

# Function to modify a single Excel workbook
def modify_workbook(filename):
    file_path = os.path.join(WORKBOOK_DIR, filename)
    try:
        workbook = openpyxl.load_workbook(file_path)

        # Create a style for date formatting
        date_style = NamedStyle(name="datetime", number_format='yyyy-mm-dd')
        
        if "datetime" not in workbook.named_styles:
            workbook.add_named_style(date_style)
    
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            print(f"Modifying sheet: {sheet_name}")

            # Step 1: Find the actual last row with data in the worksheet
            last_row = ws.max_row
            while last_row > 1 and ws.cell(row=last_row, column=1).value in (None, ""):
                last_row -= 1

            # Step 2: Determine the next date to add
            last_date = None
            for row in range(last_row, 1, -1):
                cell_value = ws.cell(row=row, column=1).value
                if isinstance(cell_value, datetime):
                    last_date = cell_value
                    break
                elif isinstance(cell_value, str):
                    try:
                        last_date = parser.parse(cell_value)
                        break
                    except ValueError:
                        continue  # Skip rows that cannot be parsed as a date

            if last_date is None:
                # If no valid date is found, set a fallback date
                last_date = datetime.now() - timedelta(days=1)

            next_date = last_date + timedelta(days=1)

            # Step 3: Identify the last non-zero NAV in column J (NAV)
            nav_column_index = 10
            last_non_zero_nav = None

            for row in range(last_row, 1, -1):
                nav_value = ws.cell(row=row, column=nav_column_index).value
                if isinstance(nav_value, (int, float)) and nav_value != 0:
                    last_non_zero_nav = nav_value
                    break

            if last_non_zero_nav is None:
                last_non_zero_nav = 100

            # Step 4: Identify existing stock symbols and quantities in columns C to G
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

            for col in range(3, 8):
                stock_symbol = ws.cell(row=stocks_row, column=col).value
                quantity = ws.cell(row=quantities_row, column=col).value
                if stock_symbol and isinstance(stock_symbol, str):
                    stocks[stock_symbol] = stock_symbol
                    quantities.append(quantity)

            # Step 5: Fetch historical stock data
            today_date = datetime.now().strftime('%Y-%m-%d')
            next_date_str = next_date.strftime('%Y-%m-%d')

            all_prices = {}
            for stock_symbol in stocks.keys():
                ticker = yf.Ticker(stock_symbol)

                try:
                    hist = ticker.history(start=next_date_str, end=today_date, interval="1d", auto_adjust=False)
                    if hist.empty:
                        print(f"No data found for {stock_symbol}. Skipping.")
                        continue

                    closing_prices = hist['Close'].tolist()
                    closing_dates = hist.index.strftime('%Y-%m-%d').tolist()
                    all_prices[stock_symbol] = (closing_dates, closing_prices)

                except Exception as e:
                    print(f"Error fetching data for {stock_symbol}: {e}")
                    continue

            # Step 6: Insert the fetched data and perform calculations
            basket_values = []
            returns = []
            nav_values = [last_non_zero_nav]

            for i in range(len(closing_dates)):
                # Convert the current date to datetime.date for comparison
                current_date = datetime.strptime(closing_dates[i], '%Y-%m-%d').date()

                # Check if the date already exists in the worksheet
                if any(ws.cell(row=r, column=1).value == current_date for r in range(2, last_row + 1)):
                    print(f"Date {current_date} already exists. Skipping.")
                    continue

                current_row = last_row + 1  # Add data to the immediate next row after the last data row
                last_row += 1

                # Insert date
                date_value = datetime.strptime(closing_dates[i], '%Y-%m-%d')
                date_cell = ws.cell(row=current_row, column=1, value=date_value)
                date_cell.number_format = 'yyyy-mm-dd'  # Apply the date style to the cell

                # Calculate basket value for the current date
                basket_value = 0
                for j, stock_symbol in enumerate(stocks.keys()):
                    price = all_prices[stock_symbol][1][i] if i < len(all_prices[stock_symbol][1]) else 0
                    quantity = quantities[j]
                    basket_value += price * quantity
                    ws.cell(row=current_row, column=3 + j, value=price)  # Insert price starting from column C

                # Insert basket value in column H
                ws.cell(row=current_row, column=8, value=basket_value)
                basket_values.append(basket_value)

                # Calculate returns and insert in column I
                ret = (basket_value - basket_values[i - 1]) / basket_values[i - 1] if i > 0 and basket_values[i - 1] != 0 else 0
                returns.append(ret)
                ws.cell(row=current_row, column=9, value=ret)

                # Calculate NAV and insert in column J
                nav = nav_values[-1] * (1 + ret)
                nav_values.append(nav)
                ws.cell(row=current_row, column=10, value=nav)

        workbook.save(file_path)

    except Exception as e:
        print(f"Error modifying {filename}: {e}")

# Function to execute git commands to add, commit, and push changes
def git_add_commit_push(modified_files):
    try:
        for filename in modified_files:
            subprocess.run(["git", "add", f"{WORKBOOK_DIR}/{filename}"], check=True)

        # Check if there are changes to commit
        status_result = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True, check=True)
        
        # If there are no changes, return without committing
        if not status_result.stdout.strip():
            st.warning("No changes to commit.")
            return

        # Git commit with a single message for all files
        commit_message = f"Updated {', '.join(modified_files)} with new data"
        subprocess.run(["git", "commit", "-m", commit_message], check=True)

        # Git push to the remote repository
        subprocess.run(["git", "push"], check=True)

    except subprocess.CalledProcessError as e:
        print(f"Error during git operation: {e}")
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
