import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import altair as alt  # For more advanced charting
import openpyxl
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Font
import yfinance as yf

# Define the directory where the workbooks are stored (relative or absolute path)
WORKBOOK_DIR = "NAV"  # Update this path to where your Excel workbooks are stored

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

# Function to modify the workbook (add new data)
def modify_workbook(file_path):
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)
        
        # Modify the workbook as per your previous code logic
        modified_workbook = modify_all_sheets(workbook)
        
        # Save the modified workbook back to the same file
        modified_workbook.save(file_path)
        
    
    except Exception as e:
        st.error(f"Error modifying workbook: {e}")

# Function to modify all sheets in the Excel file
def modify_all_sheets(workbook):
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        print(f"Modifying sheet: {sheet_name}")

        # Perform modification logic (same as your current implementation)
        last_date_cell = ws.cell(row=ws.max_row, column=1).value
        if isinstance(last_date_cell, datetime):
            last_date = last_date_cell
        else:
            last_date = datetime.now() - timedelta(days=30)
        next_date = last_date + timedelta(days=1)

        nav_column_index = 10  # Column J for NAV
        last_non_zero_nav = 100  # Default to 100 if NAV is missing

        # Fetch stock data and modify the sheet (same as your current implementation)
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

        # Fetch historical stock data
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

        # Insert the fetched data and perform calculations (same as current logic)
        current_row = ws.max_row + 1
        basket_values = []
        returns = []
        nav_values = [last_non_zero_nav]

        for i in range(len(closing_dates)):
            ws.cell(row=current_row + i, column=1, value=closing_dates[i])

            basket_value = 0
            for j, stock_symbol in enumerate(stocks.keys()):
                price = all_prices[stock_symbol][1][i] if i < len(all_prices[stock_symbol][1]) else 0
                quantity = quantities[j]
                basket_value += price * quantity
                ws.cell(row=current_row + i, column=3 + j, value=price)

            ws.cell(row=current_row + i, column=8, value=basket_value)
            basket_values.append(basket_value)

            ret = (basket_value - basket_values[i - 1]) / basket_values[i - 1] if i > 0 and basket_values[i - 1] != 0 else 0
            returns.append(ret)
            ws.cell(row=current_row + i, column=9, value=ret)

            nav = nav_values[-1] * (1 + ret)
            nav_values.append(nav)
            ws.cell(row=current_row + i, column=10, value=nav)

    # Return the modified workbook so it can be saved later
    return workbook

# Streamlit app layout and logic
def main():
    st.title("NAV Data Dashboard")

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
        # Trigger modification when the user selects a workbook or a date range
        file_path = os.path.join(WORKBOOK_DIR, selected_workbook)
        
        modify_workbook(file_path)  # Automatically modify the selected workbook

        # Load modified NAV data from the selected workbook
        nav_data = load_nav_data(file_path)

        # Check if NAV data is successfully loaded
        if not nav_data.empty:
           

            # Remove column B ('Stocks') if it exists
            nav_data = nav_data.drop(columns=['Stocks'], errors='ignore')

            # Rename column 'Unnamed: 8' to 'Returns' if it exists
            if 'Unnamed: 8' in nav_data.columns:
                nav_data = nav_data.rename(columns={'Unnamed: 8': 'Returns'})

            # Filter the data based on selected date range
            filtered_data = filter_data_by_date(nav_data, selected_range)
            filtered_data['Date'] = filtered_data['Date'].dt.date
            # Recalculate NAV to start from 100 for ranges other than '1 Day' and '5 Days'
            if selected_range not in ["1 Day", "5 Days"]:
                filtered_data = recalculate_nav(filtered_data)
                chart_column = 'Rebased NAV'
            else:
                chart_column = 'NAV'

            # Display the filtered data as a table
            

            # Generate and display the Altair line chart
            st.write("### NAV Chart")
            line_chart = alt.Chart(filtered_data).mark_line().encode(
                x='Date:T',
                y=alt.Y(f'{chart_column}:Q', scale=alt.Scale(domain=[80, filtered_data[chart_column].max()])),
                tooltip=['Date:T', f'{chart_column}:Q']
            ).properties(
                width=700,
                height=400
            )

            st.altair_chart(line_chart, use_container_width=True)
            st.write("### Data Table")
            st.dataframe(filtered_data)
if __name__ == "__main__":
    main()
