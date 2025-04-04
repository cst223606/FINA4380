import pandas as pd
import yfinance as yf

# Step 1: Get S&P 500 table from Wikipedia
url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
sp500 = pd.read_html(url)[0]

# Step 2: Clean tickers (Yahoo uses '-' instead of '.')
sp500['Symbol'] = sp500['Symbol'].str.replace('.', '-', regex=False)

# Step 3: Group tickers by sector
sector_groups = sp500.groupby('GICS Sector')

# Step 4: Create a dictionary of DataFrames: one per sector
sector_dfs = {sector: group[['Symbol']] for sector, group in sector_groups}

# Step 5: Save the tickers horizontally (starting from B1) into an Excel file with each sector as a separate sheet
with pd.ExcelWriter('sp500_data_by_sector.xlsx') as writer:
    for sector, df in sector_dfs.items():
        # Transpose the tickers to make them horizontal
        transposed_df = df.T
        
        # Write the transposed DataFrame to the Excel sheet starting at B1
        transposed_df.to_excel(writer, sheet_name=sector, index=False, header=False, startrow=0, startcol=1)

tickers = 'APPL'

# Define the start and end dates for the data (you can adjust as needed)
start_date = '2020-01-01'
end_date = '2024-12-31'

# Download the stock data for the tickers
stock_data = yf.download(tickers, start=start_date, end=end_date)

# Show the first few rows of the data
print(stock_data.head())
