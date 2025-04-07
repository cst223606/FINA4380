import pandas as pd

# Load Excel file
file_path = 'Pairs.xlsx'
price_df = pd.read_excel(file_path, sheet_name=0, parse_dates=['Date'], index_col='Date')  # Prices
pair_df = pd.read_excel(file_path, sheet_name=1)  # Cointegrated pairs

print(price_df)
print(pair_df)

spread_dict = {}

for _, row in pair_df.iterrows():
    stock_y = row['stock Y']  
    stock_x = row['stock X'] 
    beta = row['beta']
    
    if stock_y in price_df.columns and stock_x in price_df.columns:
        # Calculate spread
        spread = price_df[stock_y] - beta * price_df[stock_x]
        spread_dict[f"{stock_y}-{stock_x}"] = spread

spread_df = pd.DataFrame(spread_dict)

print(spread_df)

spread_stats = spread_df.agg(['mean', 'std'])
print(spread_stats.T)  # .T to make it more readable

zscore_df = (spread_df - spread_df.mean()) / spread_df.std()
print(zscore_df)


# Initialize variables
cash = 100000
commission = 1
shares = {}  
stock_value = {}  
account_value = {}
summary = {}
position = 0

for pair in spread_df.columns:
    position = 'none'
    shares[pair] = {'stock_y': 0, 'stock_x': 0}  # No shares at the start
    stock_value[pair] = 0  # No stock value at the start
    account_value[pair] = cash  # Initial account value is just cash
    summary[pair] = []

    zscore = zscore_df[pair]

    for date, z in zscore_df.items():

        if position == 0:
            if 1.5 < z < 2:
                stock_y, stock_x = pair.split('-')
                entranceP_y = price_df[stock_y][date]
                entranceP_x = price_df[stock_x][date]
                sharestraded_y = int(cash / (entranceP_y + commission))
                sharestraded_x = int(cash / (entranceP_x + commission))
                cash -= (entranceP_y + commission) * sharestraded_y 
                cash -= (entranceP_x + commission) * sharestraded_x
                shares[pair]['stock_y'] = sharestraded_y
                shares[pair]['stock_x'] = sharestraded_x
                position = 1
            elif -2 < z < -1.5:
                entranceP_y = price_df[pair.split('-')[0]][date] 
                entranceP_x = price_df[pair.split('-')[1]][date]
                sharestraded_y = int(cash / (entranceP_y + commission)) 
                sharestraded_x = int(cash / (entranceP_x + commission))
                cash -= (entranceP_y + commission) * sharestraded_y 
                cash -= (entranceP_x + commission) * sharestraded_x 
                shares[pair]['stock_y'] = sharestraded_y
                shares[pair]['stock_x'] = sharestraded_x
                position = 1

        elif position == 1:
            if abs(z) < 0.2:
                cash += (entranceP_y + commission) * shares[pair]['stock_y']
                cash += (entranceP_x + commission) * shares[pair]['stock_x']
                shares[pair] = {'stock_y': 0, 'stock_x': 0}
                position = 0

            elif abs(z) > 2: #Stop loss
                cash += (entranceP_y + commission) * shares[pair]['stock_y']
                cash += (entranceP_x + commission) * shares[pair]['stock_x']
                shares[pair] = {'stock_y': 0, 'stock_x': 0}
                position = 0
        summary[pair].append((date, cash, stock_value[pair], account_value[pair], shares[pair]['stock_y'], shares[pair]['stock_x']))

# Convert to DataFrame
summary_df = pd.DataFrame({
    pair: pd.Series([(date, cash, stock_val, acc_val, stock_y, stock_x) 
                     for date, cash, stock_val, acc_val, stock_y, stock_x in signal_list])
    for pair, signal_list in summary.items()
})

# Split into different columns
summary_df[['Date', 'Cash', 'Stock_Value', 'Account_Value', 'Stock_Y_Shares', 'Stock_X_Shares']] = pd.DataFrame(
    summary_df[0].to_list(), index=summary_df.index)

summary_df.drop(0, axis=1, inplace=True)





