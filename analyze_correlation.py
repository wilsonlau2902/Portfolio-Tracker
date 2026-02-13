import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def analyze_correlation():
    # Define tickers
    # Assuming 'SPYM' was a typo for 'SPY' (S&P 500), but we can check both if needed.
    # We will use SPY as the benchmark for the market.
    tickers = ['KO', 'SPY']
    
    print(f"Fetching data for {tickers}...")
    
    # Download historical data (last 10 years)
    data = yf.download(tickers, period="10y", interval="1d")
    
    # 'Adj Close' accounts for dividends and splits, which is better for calculating returns
    if 'Adj Close' in data:
        prices = data['Adj Close']
    else:
        prices = data['Close'] # Fallback if Adj Close not returned appropriately in multi-index
        
    # Calculate daily percentage returns
    returns = prices.pct_change().dropna()
    
    # 1. Calculate Correlation
    correlation = returns['KO'].corr(returns['SPY'])
    print(f"\nCorrelation between KO and SPY: {correlation:.4f}")
    
    if correlation > 0.5:
        print("Interpretation: Strong Positive Correlation (They move together)")
    elif correlation < -0.5:
        print("Interpretation: Strong Negative Correlation (They move opposite)")
    else:
        print("Interpretation: Weak or No Significant Correlation")

    # 2. Count days with opposite separate movements
    # Create a dataframe to compare signs
    moves = pd.DataFrame()
    moves['KO_Sign'] = returns['KO'].apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))
    moves['SPY_Sign'] = returns['SPY'].apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))
    
    moves['Opposite'] = (moves['KO_Sign'] * moves['SPY_Sign']) < 0
    
    opposite_days = moves['Opposite'].sum()
    total_days = len(moves)
    percentage_opposite = (opposite_days / total_days) * 100
    
    print(f"\nTotal Trading Days Analyzed: {total_days}")
    print(f"Days moving in opposite directions: {opposite_days} ({percentage_opposite:.1f}%)")
    
    # 3. Visualization
    plt.figure(figsize=(10, 6))
    sns.regplot(x=returns['SPY'], y=returns['KO'], scatter_kws={'alpha':0.5}, line_kws={'color':'red'})
    plt.title(f'Daily Returns: KO vs SPY (Correlation: {correlation:.2f})')
    plt.xlabel('SPY Daily Return')
    plt.ylabel('KO Daily Return')
    plt.grid(True)
    plt.savefig('correlation_plot.png')
    print("\nPlot saved as 'correlation_plot.png'")

if __name__ == "__main__":
    analyze_correlation()
