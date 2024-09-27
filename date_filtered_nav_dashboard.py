import streamlit as st
import pandas as pd
import os
from datetime import timedelta, datetime
import altair as alt
import openpyxl
from io import BytesIO
import yfinance as yf

# Define the directory where the workbooks are stored (locally)
WORKBOOK_DIR = "NAV"  # Ensure this path is correct for where your Excel workbooks are stored

# --- HELPER FUNCTIONS ---
# Function to list available Excel files in the specified directory (local)
def list_workbooks(directory):
    try:
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load NAV data from a workbook
def load_nav_data(file_path):
    try:
        data = pd.read_excel(file_path, sheet_name=0, usecols="A:J")
        if 'NAV' not in data.columns or 'Date' not in data.columns:
            st.error("NAV or Date column not found in the selected workbook.")
            return pd.DataFrame()

        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        data = data.sort_values(by='Date')
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
    initial_nav = filtered_data['NAV'].iloc[0]
    filtered_data['Rebased NAV'] = (filtered_data['NAV'] / initial_nav) * 100
    return filtered_data

# Function to modify all sheets in the Excel file
def modify_all_sheets(workbook, file_path):
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        last_date_cell = ws.cell(row=ws.max_row, column=1).value
        if isinstance(last_date_cell, datetime):
            last_date = last_date_cell
        else:
            last_date = datetime.now() - timedelta(days=30)
        next_date = last_date + timedelta(days=1)

        # Identify the last non-zero NAV in column J (NAV)
        nav_column_index = 10  # Column J for NAV
        last_non_zero_nav = None
        for row in range(ws.max_row, 2, -1):
            nav_value = ws.cell(row=row, column=nav_column_index).value
            if isinstance(nav_value, (int, float)) and nav_value != 0:
                last_non_zero_nav = nav_value
                break
        if last_non_zero_nav is None:
            last_non_zero_nav = 100

        # Identify existing stock symbols and quantities in columns C to G
        stocks_row = None
        quantities_row = None
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=2).value
            if cell_value == "Stocks":
                stocks_row = row
            elif cell_value == "Quantities":
                quantities_row = row

        if not stocks_row or not quantities_row:
            continue

        stocks = {}
        quantities = []
        for col in range(3, 8):  # Columns C to G
            stock_symbol = ws.cell(row=stocks_row, column=col).value
            quantity = ws.cell(row=quantities_row, column=col).value
            if stock_symbol and isinstance(stock_symbol, str):
                stocks[stock_symbol] = stock_symbol
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
                    continue
                closing_prices = hist['Close'].tolist()
                closing_dates = hist.index.strftime('%Y-%m-%d').tolist()
                all_prices[stock_symbol] = (closing_dates, closing_prices)
            except Exception as e:
                st.error(f"Error fetching data for {stock_symbol}: {e}")
                continue

        # Insert the fetched data and perform calculations
        current_row = ws.max_row + 1
        basket_values = []
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
            if i > 0:
                ret = (basket_value - basket_values[i - 1]) / basket_values[i - 1] if basket_values[i - 1] != 0 else 0
                nav = nav_values[-1] * (1 + ret)
            else:
                nav = nav_values[-1]
            basket_values.append(basket_value)
            nav_values.append(nav)
            ws.cell(row=current_row + i, column=10, value=nav)


    
    workbook.save(file_path)
         

    return workbook

# Function to save the modified Excel file locally


# --- STREAMLIT MAIN LOGIC ---
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the local NAV directory
    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    # Dropdown menus for workbook and date range selection
    selected_workbook = st.selectbox("Select a workbook", workbooks)
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    if selected_workbook:
        file_path = os.path.join(WORKBOOK_DIR, selected_workbook)
        nav_data = load_nav_data(file_path)
        if not nav_data.empty:
            st.success("Data loaded successfully!")
            filtered_data = filter_data_by_date(nav_data, selected_range)
            filtered_data['Date'] = filtered_data['Date'].dt.date

            # Recalculate NAV for ranges other than "1 Day" and "5 Days"
            if selected_range not in ["1 Day", "Max"]:
                filtered_data = recalculate_nav(filtered_data)
                chart_column = 'Rebased NAV'
            else:
                chart_column = 'NAV'

            # Display chart
            line_chart = alt.Chart(filtered_data).mark_line().encode(
                x='Date:T',
                y=alt.Y(f'{chart_column}:Q', scale=alt.Scale(domain=[80, filtered_data[chart_column].max()])),
                tooltip=['Date:T', f'{chart_column}:Q']
            ).properties(width=700, height=400)
            st.altair_chart(line_chart, use_container_width=True)

            # Display table
            st.write("### Data Table")
            st.dataframe(filtered_data.reset_index(drop=True))

            # Auto-run Excel modification when a selection is made
            workbook = openpyxl.load_workbook(file_path)
            modified_workbook = modify_all_sheets(workbook, file_path)
           

             # Refresh the dashboard after modification


if __name__ == "__main__":
    main()
