import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
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

# 2. Open your Sheet
# Note: Ensure you have shared your Google Sheet with the client_email from service_account.json
# Sheet ID from URL: https://docs.google.com/spreadsheets/d/19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESayGbTO8/edit
SHEET_ID = "19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESayGbTO8"

try:
    sheet = client.open_by_key(SHEET_ID).sheet1
except gspread.exceptions.SpreadsheetNotFound:
    print(f"Error: Spreadsheet with ID '{SHEET_ID}' not found. Make sure it is shared with the service account email.")
    exit(1)

# 3. Define your Portfolio (from your screenshots)
# Format: {Ticker: Quantity}
portfolio = {
    "AMD": 1, "AMZN": 1, "BABA": 1, "BAC": 5, 
    "EWJ": 1, "GLDM": 1, "GOOGL": 1, "INDA": 6, 
    "KO": 22, "NKE": 6, "NVDA": 2, "SPYM": 29, 
    "XAIX": 4, "XLU": 8
}

def update_portfolio():
    print("Fetching market data...")
    tickers = list(portfolio.keys())
    
    # Downloading data
    # period="5d" is safer to ensure we get at least one row of data even over weekends/holidays
    data = yf.download(tickers, period="5d")
    
    if data.empty:
        print("Error: No data fetched from yfinance.")
        return

    # Handling the data structure from yfinance
    try:
        # Access 'Close' column. If mult-level index (common in recent yfinance), this works.
        close_data = data['Close']
        
        if close_data.empty:
             print("Error: 'Close' data is empty.")
             return
             
        # Get the most recent row (last valid trading day)
        current_prices = close_data.iloc[-1]
        
    except KeyError:
        print("Error: Could not find 'Close' data in yfinance response.")
        return
    except IndexError:
         print("Error: Data index out of bounds.")
         return

    print("Updating Google Sheet...")
    rows = [["Ticker", "Qty", "Price", "Total Value"]]
    total_portfolio_value = 0
    
    for ticker, qty in portfolio.items():
        # yfinance might return NaN if data isn't available
        try:
            price = current_prices.get(ticker)
        except:
             price = None

        # Check if price is valid number (not NaN, not None)
        if price is not None and pd.notna(price):
            value = price * qty
            rows.append([ticker, qty, round(price, 2), round(value, 2)])
            total_portfolio_value += value
        else:
            rows.append([ticker, qty, "N/A", "N/A"])

    # Add a total row
    rows.append(["TOTAL", "", "", round(total_portfolio_value, 2)])
    
    # Update the sheet starting at cell A1
    try:
        # Try new gspread method first
        sheet.update(range_name='A1', values=rows)
    except TypeError:
         # Fallback for older legacy method
         sheet.update('A1', rows)

    print(f"Successfully updated 'Portfolio Tracker' at {time.strftime('%H:%M:%S')}")

if __name__ == "__main__":
    update_portfolio()
