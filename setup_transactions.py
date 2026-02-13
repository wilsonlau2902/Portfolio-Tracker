import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import time
import os

# 1. Setup Authentication
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Ensure the service account file exists
creds_file = "service_account.json"
if not os.path.exists(creds_file):
    print(f"Error: {creds_file} not found. Please ensure it is in the same directory.")
    exit(1)

creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
client = gspread.authorize(creds)

# Sheet ID from update_portfolio.py
SHEET_ID = "19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESayGbTO8"

try:
    spreadsheet = client.open_by_key(SHEET_ID)
except gspread.exceptions.SpreadsheetNotFound:
    print(f"Error: Spreadsheet with ID '{SHEET_ID}' not found.")
    exit(1)

# Initial Data provided by user
# Format: Ticker: (Qty, Price)
initial_holdings = {
    "AMD": (1, 204.0),
    "AMZN": (1, 230.5),
    "BABA": (1, 160.0),
    "BAC": (5, 55.44),
    "EWJ": (1, 84.07),
    "GLDM": (1, 100.0),
    "GOOGL": (1, 296.94),
    "INDA": (6, 56.289),
    "KO": (22, 69.629),
    "NKE": (6, 62.645),
    "NVDA": (2, 181.5),
    "SPYM": (29, 77.27),
    "XAIX": (4, 43.15),
    "XLU": (8, 44.066)
}

def setup_sheets():
    print("Setting up sheets...")

    # 1. Setup Transactions Tab
    try:
        trans_sheet = spreadsheet.worksheet("Transactions")
        print("Found existing 'Transactions' sheet.")
    except gspread.exceptions.WorksheetNotFound:
        print("Creating 'Transactions' sheet...")
        trans_sheet = spreadsheet.add_worksheet(title="Transactions", rows=100, cols=10)

    # Prepare initial data
    headers = ["Date", "Ticker", "Type", "Qty", "Price", "Total Capital"]
    rows = [headers]
    
    # Use today's date for initial setup or a placeholder
    date_str = time.strftime("%Y-%m-%d")
    
    for ticker, (qty, price) in initial_holdings.items():
        total_capital = qty * price
        rows.append([date_str, ticker, "Buy", qty, price, total_capital])

    # Clear and update Transactions
    trans_sheet.clear()
    trans_sheet.update(range_name='A1', values=rows)
    print("Populated 'Transactions' sheet with initial holdings.")

    # 2. Setup Dashboard Tab
    try:
        dash_sheet = spreadsheet.worksheet("Dashboard")
        print("Found existing 'Dashboard' sheet.")
    except gspread.exceptions.WorksheetNotFound:
        print("Creating 'Dashboard' sheet...")
        dash_sheet = spreadsheet.add_worksheet(title="Dashboard", rows=50, cols=10)
    
    # We will let update_portfolio.py populate the dashboard content, 
    # but we'll clear it here to be clean.
    dash_sheet.clear()
    dash_sheet.update(range_name='A1', values=[["Status", "Waiting for Python Update..."]])
    print("Prepared 'Dashboard' sheet.")

if __name__ == "__main__":
    setup_sheets()
