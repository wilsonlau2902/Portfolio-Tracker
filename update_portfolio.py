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
# Sheet ID from URL: https://docs.google.com/spreadsheets/d/19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESaYGbTO8/edit
SHEET_ID = "19d0B0GoNgPPYe7kxlQhGj1B8J3TsQN4u3yESaYGbTO8"

try:
    sheet = client.open_by_key(SHEET_ID)
except gspread.exceptions.SpreadsheetNotFound:
    print(f"Error: Spreadsheet with ID '{SHEET_ID}' not found. Make sure it is shared with the service account email.")
    exit(1)

# 3. Define Logic
def update_portfolio():
    print("Connecting to sheets...")
    try:
        trans_sheet = sheet.worksheet("Transactions")
        dash_sheet = sheet.worksheet("Dashboard")
    except gspread.exceptions.WorksheetNotFound:
        print("Error: 'Transactions' or 'Dashboard' worksheet not found. Please run setup_transactions.py first.")
        return

    print("Fetching transaction history...")
    # Get all records using gspread (returns list of dicts)
    records = trans_sheet.get_all_records()
    if not records:
        print("No transactions found.")
        return
        
    df_trans = pd.DataFrame(records)
    
    # Ensure numeric columns are actually numeric
    df_trans['Qty'] = pd.to_numeric(df_trans['Qty'])
    df_trans['Total Capital'] = pd.to_numeric(df_trans['Total Capital'])

    # 4. Aggregate Positions (WAC Model)
    # Filter for 'Buy' (Positive Qty) and 'Sell' (Negative Qty logic to be added later if needed)
    # For now, we assume simple accumulation
    print("Calculating Weighted Average Cost...")
    
    # Filter only Buys for now or assume Qty handles direction if we had sells
    # User specified "Buy" type, so let's stick to simple grouping
    summary = df_trans.groupby('Ticker').agg({
        'Qty': 'sum',
        'Total Capital': 'sum'
    })
    
    # Calculate Average Cost
    summary['Avg Cost'] = summary['Total Capital'] / summary['Qty']
    
    # Filter out closed positions (Qty 0)
    summary = summary[summary['Qty'] > 0]

    # 5. Fetch Real-Time Prices
    tickers = summary.index.tolist()
    print(f"Fetching market data for: {tickers}")
    
    # period="5d" to ensure we catch closing prices over weekends
    data = yf.download(tickers, period="5d", progress=False)
    
    if data.empty:
        print("Error: No data fetched from yfinance.")
        return

    try:
         # Access 'Close' column
         close_data = data['Close']
         # Get the most recent row (last valid trading day)
         current_prices = close_data.iloc[-1]
    except (KeyError, IndexError):
        print("Error processing yfinance data.")
        return

    # NEW: Fetch Sector Data
    print("Fetching sector information (this may take a moment)...")
    sector_map = {}
    
    try:
        ts = yf.Tickers(" ".join(tickers))
        for ticker in tickers:
            try:
                # Need to check if 'sector' exists in info
                t_info = ts.tickers[ticker].info
                sector_map[ticker] = t_info.get('sector', 'Other')
            except Exception:
                # Fallback if specific ticker info fails
                sector_map[ticker] = 'Other'
    except Exception as e:
        print(f"Warning: Batch sector fetch had issues: {e}. Defaulting to 'Other'.")

    print("Calculating Unrealized PnL & Sector Allocation...")
    output = []
    
    total_unrealized_pnl = 0
    total_market_value = 0
    total_cost_basis = 0
    
    # For Sector Aggregation
    sector_allocation = {}
    
    for ticker, row in summary.iterrows():
        try:
            price = current_prices.get(ticker)
        except:
             price = None

        if price is not None and pd.notna(price):
            cur_qty = row['Qty']
            avg_cost = row['Avg Cost']
            
            market_value = price * cur_qty
            cost_basis = row['Total Capital']
            
            unrealized_pnl = market_value - cost_basis
            pnl_pct = (unrealized_pnl / cost_basis) * 100
            
            # Get Sector
            sector = sector_map.get(ticker, 'Other')
            
            # Update Sector Totals
            sector_allocation[sector] = sector_allocation.get(sector, 0) + market_value
            
            output.append([
                ticker, 
                cur_qty, 
                round(avg_cost, 2), 
                round(price, 2), 
                round(market_value, 2), 
                sector, # Added Sector Column
                round(unrealized_pnl, 2), 
                f"{pnl_pct:.2f}%"
            ])
            
            total_unrealized_pnl += unrealized_pnl
            total_market_value += market_value
            total_cost_basis += cost_basis
        else:
            sector = sector_map.get(ticker, 'Other')
            output.append([ticker, row['Qty'], round(row['Avg Cost'], 2), "N/A", "N/A", sector, "N/A", "N/A"])

    # 6. Push to Dashboard
    headers = ["Ticker", "Qty", "Avg Cost", "Live Price", "Market Value", "Sector", "Unrealized PnL", "Unrealized %"]
    
    # Sort output by Ticker
    output.sort(key=lambda x: x[0])
    
    # Calculate Sector Percentages
    sector_output = []
    for sec, val in sector_allocation.items():
        pct = (val / total_market_value) * 100 if total_market_value > 0 else 0
        sector_output.append([sec, round(val, 2), f"{pct:.2f}%"])
    
    # Sort sectors by value (descending)
    sector_output.sort(key=lambda x: x[1], reverse=True)
    sector_headers = ["Sector", "Total Value", "% Alloc"]

    
    total_gain_pct = (total_unrealized_pnl / total_cost_basis * 100) if total_cost_basis > 0 else 0
    
    print("Updating Dashboard...")
    dash_sheet.clear()
    
    # Ensure sheet has enough columns (M is col 13, let's set to 20 to be safe)
    if dash_sheet.col_count < 20:
        dash_sheet.resize(cols=20)
    
    # Update Main Table (A:H)
    dash_sheet.update(range_name='A1', values=[headers] + output)
    
    # Update Portfolio Summary (J:K)
    summary_data = [
        ["PORTFOLIO SUMMARY", ""],
        ["Total Market Value", round(total_market_value, 2)],
        ["Total Cost Basis", round(total_cost_basis, 2)],
        ["Total Unrealized PnL", round(total_unrealized_pnl, 2)],
        ["Total Return %", f"{total_gain_pct:.2f}%"],
        ["Last Updated", time.strftime('%Y-%m-%d %H:%M:%S')]
    ]
    dash_sheet.update(range_name='J1', values=summary_data)
    
    # Update Sector Allocation Table (M:O)
    dash_sheet.update(range_name='M1', values=[sector_headers] + sector_output)

    
    print(f"Dashboard fully synced at {time.strftime('%H:%M:%S')}")

if __name__ == "__main__":
    update_portfolio()
