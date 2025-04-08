import pandas as pd
import matplotlib.pyplot as plt

# Load pair data
file_path = '../../new pairs.xlsx'
pair_df = pd.read_excel(file_path)
print(pair_df.head())

# Extract needed tickers from 'Stock X' and 'Stock Y'
tickers_needed = set(pair_df['Stock Y']).union(set(pair_df['Stock X']))

# Initialize dictionary to store stock prices
stock_prices = {}
added_tickers = set()

# Load stock prices from the sector data file
sector_file = '../../sp500_data_by_sector.xlsx'
xls = pd.ExcelFile(sector_file)

# Function to clean price values (remove dollar signs, commas, and convert to numeric)
def clean_price(price):
    if isinstance(price, str):
        price = price.replace('$', '').replace(',', '')  # Remove $ and commas
    return pd.to_numeric(price, errors='coerce')  # Convert to float, replace invalid values with NaN

# Load and clean stock price data
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if 'Date' column exists and set it as index
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])
        df.set_index('Date', inplace=True)

    # Add stock prices to dictionary and clean them
    for col in df.columns:
        if col in tickers_needed and col not in added_tickers:
            # Clean the stock prices before adding to the dictionary
            df[col] = df[col].apply(clean_price)  # Clean prices for this stock
            stock_prices[col] = df[col]
            added_tickers.add(col)

# Combine stock prices into DataFrame and sort by date
price_df = pd.DataFrame(stock_prices)
price_df = price_df.sort_index(ascending=True)
print(price_df.head())

# Create spread based on pairs in pair_df
spread_dict = {}
for _, row in pair_df.iterrows():
    stock_y = row['Stock Y']
    stock_x = row['Stock X']
    beta = row['Beta']
    
    if stock_y in price_df.columns and stock_x in price_df.columns:
        spread = price_df[stock_y] - beta * price_df[stock_x]
        spread_dict[f"{stock_y}-{stock_x}"] = spread

spread_df = pd.DataFrame(spread_dict)
print(spread_df.head())

# Calculate z-scores for each pair's spread
zscore_df = (spread_df - spread_df.mean()) / spread_df.std()
print(zscore_df.head())

# Trading parameters
commission = 1
initial_cash = 1000

# Initialize storage for trading data
cash = {}
shares = {}
stock_value = {}
summary = {}
positions = {}
sub_value = {}

# Execute trades for each pair
for _, row in pair_df.iterrows():
    stock_y = row['Stock Y']
    stock_x = row['Stock X']
    beta = row['Beta']
    
    pair = f"{stock_y}-{stock_x}"
    cash[pair] = initial_cash
    shares[pair] = {'stock_y': 0, 'stock_x': 0}
    stock_value[pair] = 0
    summary[pair] = []
    positions[pair] = 0
    zscore = zscore_df[pair]
    
    for date, z in zscore.items():
        price_y = price_df.at[date, stock_y]
        price_x = price_df.at[date, stock_x]

        # Entry signal: 
        if positions[pair] == 0:
            # Long Y, Short X
            if 1.5 < z < 2:
                shares_y = int(cash[pair] / (2 * (price_y + commission)))
                shares_x = int(cash[pair] / (2 * (price_x + commission)))
                cash[pair] -= (price_y + commission) * shares_y
                cash[pair] += (price_x + commission) * shares_x
                shares[pair]['stock_y'] = shares_y
                shares[pair]['stock_x'] = -shares_x
                positions[pair] = 1
            # Long X, Short Y
            elif -2 < z < -1.5:
                shares_y = int(cash[pair] / (2 * (price_y + commission)))
                shares_x = int(cash[pair] / (2 * (price_x + commission)))
                cash[pair] += (price_y - commission) * shares_y
                cash[pair] -= (price_x - commission) * shares_x
                shares[pair]['stock_y'] = -shares_y
                shares[pair]['stock_x'] = shares_x
                positions[pair] = 1

        # Exit condition: Convergence or stop-loss
        elif positions[pair] != 0:
            if abs(z) < 0.2 or abs(z) > 2:
                shares_y = shares[pair]['stock_y']
                shares_x = shares[pair]['stock_x']
                cash[pair] += shares_y * (price_y + commission if shares_y > 0 else price_y + commission)
                cash[pair] += shares_x * (price_x + commission if shares_x > 0 else price_x + commission)
                shares[pair] = {'stock_y': 0, 'stock_x': 0}
                positions[pair] = 0

        # Update portfolio value
        stock_value[pair] = (shares[pair]['stock_y'] * price_y + shares[pair]['stock_x'] * price_x) if positions[pair] != 0 else 0
        sub_value[pair] = cash[pair] + stock_value[pair]

        summary[pair].append({
            'Date': date,
            'Cash': cash[pair],
            'Net Stock_Value': stock_value[pair],
            'sub_value': sub_value[pair],
            'Stock_Y_Shares': shares[pair]['stock_y'],
            'Stock_X_Shares': shares[pair]['stock_x'],
            'Z-Score': z,
            'Position': positions[pair]
        })

# Calculate total portfolio value
account_value_df = pd.DataFrame({
    pair: pd.DataFrame(data).set_index('Date')['sub_value']
    for pair, data in summary.items()
})
account_value_df['Total_Account_Value'] = account_value_df.sum(axis=1)
print(account_value_df.head())
# Output to Excel
#with pd.ExcelWriter('Trading_Summary_Output.xlsx') as writer:
   # account_value_df.to_excel(writer, sheet_name='Portfolio_Account_Value')

# Plot total account value
plt.figure(figsize=(12, 6))
plt.plot(account_value_df.index, account_value_df['Total_Account_Value'], label='Total Portfolio Value', color='blue')
plt.title('Total Portfolio Account Value Over Time')
plt.xlabel('Date')
plt.ylabel('Account Value')
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.savefig('Total_Account_Value_Plot.png')
#plt.show()
