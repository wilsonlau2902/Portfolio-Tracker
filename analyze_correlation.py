import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import pandas as pd
import time
import os
import datetime

# 1. Setup Authentication (Reusing your existing setup)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_file = "service_account.json"
SHEET_ID = "19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESaYGbTO8"

def get_client():
    if not os.path.exists(creds_file):
        print(f"Error: {creds_file} not found.")
        return None
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
    return gspread.authorize(creds)

def setup_correlation_sheet(sheet):
    """Creates the Correlation tab layout if it doesn't exist or is empty."""
    try:
        ws = sheet.worksheet("Correlation Analysis")
    except gspread.exceptions.WorksheetNotFound:
        print("Creating 'Correlation Analysis' tab...")
        ws = sheet.add_worksheet(title="Correlation Analysis", rows=100, cols=20)
        
        # Initial Layout
        ws.update(range_name='A1', values=[
            ["CONFIGURATION", ""],
            ["Time Period (1y, 2y, 3y, 5y, 10y)", "1y"],
            ["Tickers to Analyze (One per row)", "Last Updated:"],
            ["AMD", ""], 
            ["NVDA", ""],
            ["GOOGL", ""],
            ["MSFT", ""],
            ["SPY", ""]
        ])
    return ws

def run_correlation_analysis():
    client = get_client()
    if not client: return

    try:
        sheet = client.open_by_key(SHEET_ID)
    except Exception as e:
        print(f"Error opening sheet: {e}")
        return

    # 1. Get or Create Tab
    ws = setup_correlation_sheet(sheet)
    
    print("\nReading configuration from 'Correlation Analysis' tab...")
    
    # Read Period (Cell B2) of 'Correlation Analysis'
    period_raw = ws.acell('B2').value
    # Normalize input: remove spaces, convert to lowercase
    period = str(period_raw).strip().lower() if period_raw else "1y"
    
    valid_periods = ['1y', '2y', '5y', '10y', 'ytd', 'max']
    
    # Allow '3y' even if possibly not standard in some contexts, yfinance supports it.
    if period not in valid_periods and 'y' not in period:
         print(f"Warning: value '{period_raw}' might be invalid. Defaulting to '1y'.")
         period = "1y"
         ws.update('B2', "1y")

    print(f"Configuration Period: {period}")
    
    # Read Tickers (Column A, starting A4)
    # Read all values in Col A
    col_a = ws.col_values(1)
    
    # Extract tickers starting from row 4
    # The list index is 0-based, so Row 4 is index 3
    if len(col_a) > 3:
        tickers = [t.strip().upper() for t in col_a[3:] if t.strip()]
    else:
        tickers = []

    # Filter duplicates while maintaining order
    tickers = list(dict.fromkeys(tickers))

    if len(tickers) < 2:
        print("Not enough tickers found in Column A (rows 4+). Please add at least 2 tickers in the sheet.")
        return

    print(f"Analyzing {len(tickers)} tickers over {period}...")

    # 2. Fetch Data
    try:
        # progress=False suppresses the progress bar string output
        data = yf.download(tickers, period=period, progress=False)
    except Exception as e:
        print(f"Error fetching data: {e}")
        return

    if data.empty:
        print("No data fetched. Check ticker symbols.")
        return

    try:
        # Handle yfinance multi-index columns if needed
        # 'Close' usually returns a DataFrame where columns are Tickers
        if 'Close' in data:
            prices = data['Close']
        else:
            prices = data
        
        # 1. Drop columns (tickers) that have NO data at all (avoid failure if one ticker is bad)
        prices = prices.dropna(axis=1, how='all')
        
        if prices.empty:
             print("All tickers returned empty data.")
             return

        # 2. Calculate daily percentage returns
        returns = prices.pct_change()
        
        # 3. Allow partial data (pairwise correlation handles it)
        # We only drop rows that are completely empty across ALL tickers (e.g. holidays)
        returns = returns.dropna(how='all') 
        
        print(f"Proceeding with {len(returns)} rows of data...")

    except Exception as e:
        print(f"Error processing data structure: {e}")
        return

    # 3. Calculate Correlation
    # min_periods=1 allows calculation even if overlaps are sparse
    corr_matrix = returns.corr()

    # 4. Output to Sheet
    # We'll place the matrix starting at Cell E1
    
    print("Writing heatmap data to sheet...")
    
    # Clean non-serializable/NaNs if any (using fillna)
    corr_matrix = corr_matrix.fillna(0)
    
    # Prepare output list
    # Row 1: [Space, T1, T2, T3...] (Header)
    header_row = ["Correlation Matrix"] + corr_matrix.columns.tolist()
    
    output_data = [header_row]
    for index, row in corr_matrix.iterrows():
        # [Ticker, Val1, Val2, Val3...]
        row_data = [index] + row.tolist()
        # Round logic for display
        row_data = [x if isinstance(x, str) else round(x, 2) for x in row_data]
        output_data.append(row_data)

    # Clear previous results area (E1 to AZ100)
    try:
        ws.batch_clear(["E1:AZ100"])
        ws.update(range_name='E1', values=output_data)
        
        # Write Timestamp to C3
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Ensure C3 has a value
        try:
             ws.update(range_name='C3', values=[[f"Last Updated: {current_time}"]])
        except:
             pass
             
        print("Success! Data updated.")
        
    except Exception as e:
        print(f"Error updating sheet: {e}")

if __name__ == "__main__":
    run_correlation_analysis()
