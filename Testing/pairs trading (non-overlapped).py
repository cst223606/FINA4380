import pandas as pd
import numpy as np

file_path = '../new pairs.xlsx'
pair_df = pd.read_excel(file_path)
pair_df = pair_df[pair_df['Beta'] > 0.65]
pair_df = pair_df[pair_df['Beta'] < 1.5]
pair_df = pair_df.sort_values(by='ADF P-value', ascending=True).reset_index(drop=True)
used_stocks = set()
filtered_rows = []

for _, row in pair_df.iterrows():
    stock_x = row['Stock X']
    stock_y = row['Stock Y']
    
    # If neither stock has been used, keep the pair
    if stock_x not in used_stocks and stock_y not in used_stocks:
        used_stocks.update([stock_x, stock_y])
        filtered_rows.append(row)

# Create new DataFrame of filtered rows
filtered_pair_df = pd.DataFrame(filtered_rows).reset_index(drop=True)
print(filtered_pair_df)
#target_pair = 'AFL-EG'

# Split the pair name into stock_y and stock_x
#stock_x, stock_y = target_pair.split('-')

# Filter just this pair
#filtered_pair_df = pair_df[(pair_df['Stock Y'] == stock_y) & (pair_df['Stock X'] == stock_x)]

# Extract needed tickers from 'Stock X' and 'Stock Y'
tickers_needed = set(filtered_pair_df['Stock Y']).union(set(filtered_pair_df['Stock X']))

# Initialize dictionary to store stock prices
stock_prices = {}
added_tickers = set()

# Load stock prices from the sector data file
sector_file = '../sp500_data_by_sector.xlsx'
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

spread_dict = {}
for _, row in pair_df.iterrows():
    stock_y = row['Stock Y']
    stock_x = row['Stock X']
    beta = row['Beta']
    
    if stock_y in price_df.columns and stock_x in price_df.columns:
        spread = price_df[stock_y] - beta * price_df[stock_x]
        spread_dict[f"{stock_y}-{stock_x}"] = spread

spread_df = pd.DataFrame(spread_dict)
zscore_df = (spread_df - spread_df.mean()) / spread_df.std()
#print(spread_df)
print(zscore_df)
zscore_df.to_excel('Zscore_Output.xlsx', sheet_name='Z-Scores')



# Parameters
initial_capital = 100000000
n_pairs = len(filtered_pair_df)
capital_per_pair = initial_capital / n_pairs
commission = 0.005
short_interest_rate = 0.1 / 365
cash_interest_rate = 0.04 / 365
daily_short_interest = 0
daily_cash_interest = 0
reserved_cash = 0
n = 0
# Results
account_value_df = pd.DataFrame(index=zscore_df.index)
pair_values = {}

# Loop through each pair
for _, row in filtered_pair_df.iterrows():
    x = row['Stock X']
    y = row['Stock Y']
    beta = row['Beta']
    pair_name = f"{y}-{x}"
    
    zscores = zscore_df[pair_name]
    prices_x = price_df[x]
    prices_y = price_df[y]
    
    position = 0  # 0 or 1 or -1
    units_x = units_y = 0
    stock_value = 0
    values = []
    trade = 0
    cash = capital_per_pair
    residual_cash = 0
    total_cash = cash + reserved_cash

    for date in zscores.index:
        price_x = prices_x.loc[date]
        price_y = prices_y.loc[date]
        z = zscores.loc[date]
        
        # Entry logic
        if position == 0:
            if 1.7 < z < 2:
                reserved_cash = total_cash * 0.05
                cash = total_cash * 0.95
                # Short 1/Beta Y, Long X
                x_entranceP = price_x
                y_entranceP = price_y
                units_x = int(cash / (x_entranceP + commission))
                units_y = -int(units_x * 1 / beta)
                while (cash - (x_entranceP + commission) * units_x + units_y * commission) < 0 and units_x > 0 :
                    units_x = units_x - 1
                    units_y = -int(units_x * 1 / beta)
                residual_cash = cash - (x_entranceP + commission) * units_x + units_y * commission
                cash = residual_cash - y_entranceP * units_y
                position = -1
                trade = trade + 1

            elif -2 < z < -1.7 :
                reserved_cash = total_cash * 0.05
                cash = total_cash * 0.95
                # Long Y, Short Beta X
                x_entranceP = price_x
                y_entranceP = price_y
                units_y = int(cash / (y_entranceP + commission))
                units_x = -int(units_y * beta)
                while (cash - (y_entranceP + commission) * units_y + units_x * commission) < 0 and units_y > 0 :
                    units_y = units_y - 1
                    units_x = -int(units_y * beta)
                residual_cash = cash - (y_entranceP + commission) * units_y + units_x * commission
                cash = residual_cash - x_entranceP * units_x
                position = 1
                trade = trade + 1
        
        # Exit logic
        elif position != 0:
            if abs(z) < 0.3 or abs(z) > 2.3 or reserved_cash < daily_short_interest :
                # Close position
                if position == 1 :
                    cash = residual_cash - price_x * units_x
                    cash = cash + units_y * price_y + units_x * price_x - commission * (units_y - units_x)
                elif position == -1 :
                    cash = residual_cash - price_y * units_y
                    cash = cash + units_y * price_y + units_x * price_x - commission * (units_x - units_y)
                units_x = units_y = 0
                position = 0
                trade = trade + 1

        # Compute daily value
        if position == 1:
            cash = residual_cash - price_x * units_x
            daily_short_interest = short_interest_rate * units_x * x_entranceP
            daily_cash_interest = cash_interest_rate * (residual_cash + reserved_cash)
        elif position == -1 :
            cash = residual_cash - price_y * units_y
            daily_short_interest = short_interest_rate * units_y * y_entranceP
            daily_cash_interest = cash_interest_rate * (residual_cash + reserved_cash)
        elif position == 0 :
            daily_cash_interest = cash_interest_rate * (cash + reserved_cash)
            daily_short_interest = 0
        reserved_cash = reserved_cash + daily_cash_interest + daily_short_interest
        stock_value = units_x * price_x + units_y * price_y
        total_cash = cash + reserved_cash
        value = total_cash + stock_value
        values.append(value)

    # Store results
    pair_values[pair_name] = values
    print(pair_name)
    print(reserved_cash)
    print(value)

# Combine into DataFrame and sum across all pairs
pair_values_df = pd.DataFrame(pair_values, index=zscore_df.index)
account_value_df['Account Value'] = pair_values_df.sum(axis=1)
account_value_df.to_excel('Account Values (Non-Overlapped).xlsx')
#print(account_value_df)
# Plotting
import matplotlib.pyplot as plt

account_value_df.plot(figsize=(12, 6), title='Total Account Value Over Time (Non-Overlapped)')
plt.xlabel("Date")
plt.ylabel("Account Value")
plt.grid(True)
plt.savefig('Total_Account_Value_Plot_Non-overlapped.png')
plt.show()