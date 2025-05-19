import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

# Get user input for the ETF symbol
etf_symbol = input("Enter the ETF symbol: ")
etf = yf.Ticker(etf_symbol)

# Get today's date and the correct reference dates
today = datetime.today().strftime('%Y-%m-%d')
start_of_year = "2024-12-31"  # Last trading day of previous year for YTD return
one_year_ago = (datetime.today() - timedelta(days=365)).strftime('%Y-%m-%d')  # Exact one-year ago date

# Retrieve historical data
historical_data = etf.history(start=one_year_ago, end=today)

# Ensure correct adjusted closing prices
if not historical_data.empty:
    # YTD Return: Start of year vs today
    start_price_ytd = historical_data.loc[start_of_year].get("Adj Close", historical_data.loc[start_of_year]["Close"]) if start_of_year in historical_data.index else historical_data.iloc[0].get("Adj Close", historical_data.iloc[0]["Close"])
    
    # One-Year Return: One year ago vs today
    start_price_one_year = historical_data.loc[one_year_ago].get("Adj Close", historical_data.loc[one_year_ago]["Close"]) if one_year_ago in historical_data.index else historical_data.iloc[0].get("Adj Close", historical_data.iloc[0]["Close"])
    
    current_price = historical_data.iloc[-1].get("Adj Close", historical_data.iloc[-1]["Close"])  # Latest adjusted close
    
    ytd_return = ((current_price - start_price_ytd) / start_price_ytd) * 100
    one_year_return = ((current_price - start_price_one_year) / start_price_one_year) * 100
else:
    ytd_return, one_year_return = "N/A", "N/A"  # Handle missing data

# Organize data into a structured dictionary
data = {
    "Fund Name": [etf.info.get("longName", "N/A")],
    "Opening Price": [etf.info.get("open", "N/A")],
    "Yield %": [etf.info.get("yield", "N/A") * 100],
    "PE Ratio": [etf.info.get("trailingPE", "N/A")],
    "52 Week Range": [f'{etf.info.get("fiftyTwoWeekLow", "N/A")} - {etf.info.get("fiftyTwoWeekHigh", "N/A")}'],
    "YTD Return %": [round(ytd_return, 2)],  # Uses Dec 31, 2024 price
    "One-Year Return %": [round(one_year_return, 2)],  # Uses May 16, 2024 price
    "Beta": [etf.info.get("beta3Year", "N/A")]
    
}

# Convert the dictionary to a pandas DataFrame
df = pd.DataFrame(data)

# Save to Excel file
df.to_excel("ETF_Data.xlsx", index=False)
print("✅ Data saved to ETF_Data.xlsx successfully!")



