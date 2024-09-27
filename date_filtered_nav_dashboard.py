import streamlit as st
import pandas as pd
import os
from datetime import timedelta
import altair as alt  # For more advanced charting
import requests
import base64
from io import BytesIO
import openpyxl
from datetime import datetime, timedelta
import yfinance as yf

# Define the directory where the workbooks are stored
WORKBOOK_DIR = "NAV"  # Update this path to where your Excel workbooks are stored

# GitHub credentials and repository details
GITHUB_TOKEN = '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'  # Replace with your GitHub PAT
GITHUB_REPO = 'anuj1963/NAV-'  # Ensure this is the correct repository name
GITHUB_BRANCH = 'master'
NAV_FOLDER = 'NAV'  # Folder name in your GitHub repository where Excel files are stored

# Base URL for GitHub API and Raw file download
BASE_API_URL = f'https://api.github.com/repos/{GITHUB_REPO}/contents/{NAV_FOLDER}'
BASE_RAW_URL = f'https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{NAV_FOLDER}'

# Headers for GitHub API authentication
headers = {
    'Authorization': f'token {GITHUB_TOKEN}',
    'Accept': 'application/vnd.github.v3+json',
}

# Function to list available Excel files in the specified directory
def list_workbooks(directory):
    try:
        files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
        return files
    except FileNotFoundError:
        st.error("Directory not found. Please ensure the specified directory exists.")
        return []

# Function to load NAV data from the selected workbook
def load_nav_data(file_path):
    try:
        data = pd.read_excel(file_path, sheet_name=0, usecols="A:J")  # Load columns A-J
        if 'NAV' not in data.columns or 'Date' not in data.columns:
            st.error("NAV or Date column not found in the selected workbook.")
            return pd.DataFrame()

        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        data = data.sort_values(by='Date')  # Sort data by Date
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
    else:
        return data

# Function to recalculate NAV starting from 100
def recalculate_nav(filtered_data):
    initial_nav = filtered_data['NAV'].iloc[0]
    filtered_data['Rebased NAV'] = (filtered_data['NAV'] / initial_nav) * 100
    return filtered_data

# Function to modify the Excel file from GitHub and update NAV
def modify_excel_file_from_github(filename):
    # Download the Excel file from GitHub
    file_url = f'{BASE_RAW_URL}/{filename}'
    response = requests.get(file_url)
    if response.status_code != 200:
        st.error(f"Error downloading {filename} from GitHub: {response.status_code}")
        return

    excel_file = BytesIO(response.content)
    workbook = openpyxl.load_workbook(excel_file)

    # Modify the sheets in the workbook (you can insert your logic here)
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        # Add logic to modify workbook (e.g., update NAV)
        ws['A1'].value = "Modified"  # Simple example modification

    # Save the modified workbook to memory
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Upload the modified Excel file back to GitHub
    sha = get_file_sha(filename)  # Get SHA of the existing file
    upload_excel_to_github(filename, output, sha)

# Function to get the SHA of the current version of the file from GitHub
def get_file_sha(filename):
    file_url = f'{BASE_API_URL}/{filename}'
    response = requests.get(file_url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Error fetching file metadata for {filename}: {response.status_code}")

    file_metadata = response.json()
    return file_metadata['sha']

# Function to upload the modified Excel file to GitHub
def upload_excel_to_github(filename, updated_content, sha):
    encoded_content = base64.b64encode(updated_content.read()).decode('utf-8')
    commit_message = f"Updated {filename} with new data"
    data = {
        "message": commit_message,
        "content": encoded_content,
        "sha": sha,
        "branch": GITHUB_BRANCH
    }
    upload_url = f'{BASE_API_URL}/{filename}'
    response = requests.put(upload_url, headers=headers, json=data)
    if response.status_code == 200:
        st.success(f"Successfully updated {filename} on GitHub.")
    else:
        st.error(f"Failed to update {filename} on GitHub: {response.status_code}")

# Streamlit app layout and logic
def main():
    st.title("NAV Data Dashboard")

    # List available workbooks in the directory
    workbooks = list_workbooks(WORKBOOK_DIR)
    if not workbooks:
        st.error("No Excel workbooks found in the specified directory.")
        return

    # Dropdown to select a workbook
    selected_workbook = st.selectbox("Select a workbook", workbooks)

    # Date range options
    date_ranges = ["1 Day", "5 Days", "1 Month", "6 Months", "1 Year", "Max"]
    selected_range = st.selectbox("Select Date Range", date_ranges)

    # Trigger modification of the Excel file as soon as a selection is made
    if selected_workbook and selected_range:
        modify_excel_file_from_github(selected_workbook)
        st.experimental_rerun()  # Refresh the app after modification

    if selected_workbook:
        nav_data = load_nav_data(os.path.join(WORKBOOK_DIR, selected_workbook))
        if not nav_data.empty:
            st.success("Data loaded successfully!")
            filtered_data = filter_data_by_date(nav_data, selected_range)

            filtered_data['Date'] = filtered_data['Date'].dt.date
            if selected_range not in ["1 Day", "5 Days"]:
                filtered_data = recalculate_nav(filtered_data)
                chart_column = 'Rebased NAV'
            else:
                chart_column = 'NAV'

            line_chart = alt.Chart(filtered_data).mark_line().encode(
                x='Date:T',
                y=alt.Y(f'{chart_column}:Q', scale=alt.Scale(domain=[80, filtered_data[chart_column].max()])),
                tooltip=['Date:T', f'{chart_column}:Q']
            ).properties(width=700, height=400)

            st.altair_chart(line_chart, use_container_width=True)
            st.write("### Data Table")
            st.dataframe(filtered_data.reset_index(drop=True))

if __name__ == "__main__":
    main()
