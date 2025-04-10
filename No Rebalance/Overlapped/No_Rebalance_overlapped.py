import pandas as pd
import matplotlib.pyplot as plt

# Load pair datacd 
file_path = '../../new pairs.xlsx'
pair_df = pd.read_excel(file_path)
#print(pair_df.head())
num_pairs = len(pair_df)
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

latest_date = price_df.index.max()
three_years_ago = latest_date - pd.DateOffset(years=3)
price_df = price_df[price_df.index >= three_years_ago]
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
#print(spread_df.head())

# Calculate z-scores for each pair's spread
zscore_df = (spread_df - spread_df.mean()) / spread_df.std()
#zscore_df.to_excel('Zscore_Output.xlsx', sheet_name='Z-Scores')

# Initialize storage for trading data
cash = {}
shares = {}
stock_value = {}
summary = {}
positions = {}
sub_value = {}
short_interest_rate = 0.1/365 #daily short interest rate
cash_interest_rate = 0.04/365 #daily cash interest rate
short_interest = 0
cash_interest = 0
total_cash = 1000000000
# Execute trades for each pair
for _, row in pair_df.iterrows():
    # Trading parameters
    commission = 0.005
    initial_cash = total_cash / num_pairs
    position_cash = initial_cash * 0.95
    initial_cash = initial_cash - position_cash
    residual_cash = 0
    shares_y = 0
    shares_x = 0
    residual_cash = 0
    y_entranceP = 0
    x_entranceP = 0

    stock_y = row['Stock Y']
    stock_x = row['Stock X']
    beta = row['Beta']
    
    pair = f"{stock_y}-{stock_x}"
    cash[pair] = position_cash
    shares[pair] = {'stock_y': 0, 'stock_x': 0}
    stock_value[pair] = 0
    summary[pair] = []
    positions[pair] = 0
    zscore = zscore_df[pair]

    for date, z in zscore.items():
        price_y = price_df.at[date, stock_y]
        price_x = price_df.at[date, stock_x]
        residual_cash = cash[pair]
        # Entry signal: 
        if positions[pair] == 0:
            # Long Y, Short X
            if 2 < z < 2.3:
                shares_y = int(cash[pair] / ((price_y + commission)))
                shares_x = -int(shares_y * beta)
                y_entranceP = price_y
                x_entranceP = price_x
                if (cash[pair] - (y_entranceP + commission) * shares_y + shares_x * commission) < 0 :
                    shares_y = shares_y - 1
                    shares_x = -int(shares_y * beta)
                residual_cash = cash[pair] - (y_entranceP + commission) * shares_y + shares_x * commission
                cash[pair] =  residual_cash - shares_x * x_entranceP #shares_x < 0
                shares[pair]['stock_y'] = shares_y
                shares[pair]['stock_x'] = shares_x
                positions[pair] = 1
            # Long X, Short Y
            elif -2.3 < z < -2:
                shares_x = int(cash[pair] / ((price_x + commission)))
                shares_y = -int(shares_x * beta)
                x_entranceP = price_x
                y_entranceP = price_y
                if (cash[pair] - (x_entranceP + commission) * shares_x + shares_y * commission) < 0 :
                    shares_x = shares_x - 1
                    shares_y = -int(shares_x * beta)
                residual_cash = cash[pair] - (x_entranceP + commission) * shares_x + shares_y * commission
                cash[pair] =  residual_cash - shares_y * y_entranceP #shares_x < 0
                shares[pair]['stock_y'] = shares_y
                shares[pair]['stock_x'] = shares_x
                positions[pair] = -1
            
        # Exit condition: Convergence or stop-loss
        elif positions[pair] != 0:
            if abs(z) < 0.5 or abs(z) > 2.5:
                y_exitP = price_y
                x_exitP = price_x
                #Sell Y, Buy back X
                if positions[pair] == 1: 
                    cash[pair] = cash[pair] + shares_y * (y_exitP - commission) - shares_x * (commission - x_exitP)
                #Sell X, Buy back Y
                elif positions[pair] == -1:
                    cash[pair] = cash[pair] + shares_x * (x_exitP - commission) - shares_y * (commission - y_exitP)
                stock_value[pair] = 0
                shares[pair] = {'stock_y': 0, 'stock_x': 0}
                positions[pair] = 0

        # Update portfolio value
        if positions[pair] != 0:
            stock_value[pair] = shares_x * price_x + shares_y * price_y # daily net stock value
            if positions[pair] == 1:
                short_interest = short_interest_rate * shares_x * x_entranceP # shares_x < 0
            elif positions[pair] == -1:
                short_interest = short_interest_rate * shares_y * y_entranceP # shares_y < 0
            cash_interest = (initial_cash + residual_cash) * cash_interest_rate
            initial_cash = initial_cash + cash_interest + short_interest
            sub_value[pair] = cash[pair] + stock_value[pair] + initial_cash
        else :
            cash_interest = (initial_cash + cash[pair]) * cash_interest_rate
            initial_cash =  initial_cash + cash_interest #residual_cash included in cash[pair]
            sub_value[pair] = cash[pair] + initial_cash
        test = 2
        if positions[pair] == 1 or positions[pair] == -1:
            test = 1
        if date == pd.Timestamp('2025-04-04') :
            print(sub_value[pair])

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
#print(sub_value)
# Calculate total portfolio value
account_value_df = pd.DataFrame({
    pair: pd.DataFrame(data).set_index('Date')['sub_value']
    for pair, data in summary.items()
})
account_value_df['Total_Account_Value'] = account_value_df.sum(axis=1)

# Output to Excel
with pd.ExcelWriter('Trading_Summary_Output.xlsx') as writer:
    account_value_df.to_excel(writer, sheet_name='Portfolio_Account_Value')

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
