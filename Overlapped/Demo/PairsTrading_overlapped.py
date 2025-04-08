import pandas as pd
import matplotlib.pyplot as plt


file_path = '../../Pairs.xlsx'
price_df = pd.read_excel(file_path, sheet_name=0, parse_dates=['Date'], index_col='Date')  # Prices
pair_df = pd.read_excel(file_path, sheet_name=1)  # Cointegrated pairs

spread_dict = {}
for _, row in pair_df.iterrows():
    stock_y = row['stock Y']  
    stock_x = row['stock X'] 
    beta = row['beta']
    
    if stock_y in price_df.columns and stock_x in price_df.columns:
        spread = price_df[stock_y] - beta * price_df[stock_x]
        spread_dict[f"{stock_y}-{stock_x}"] = spread

spread_df = pd.DataFrame(spread_dict)

zscore_df = (spread_df - spread_df.mean()) / spread_df.std()

commission = 1
initial_cash = 100000

# Initialize storage
cash = {}
shares = {}
stock_value = {}
summary = {}
positions = {}
sub_value = {}

for pair in spread_df.columns:
    cash[pair] = initial_cash
    shares[pair] = {'stock_y': 0, 'stock_x': 0}
    stock_value[pair] = 0
    summary[pair] = []
    positions[pair] = 0  
    zscore = zscore_df[pair]
    stock_y, stock_x = pair.split('-')

    for date, z in zscore.items():
        price_y = price_df.at[date, stock_y]
        price_x = price_df.at[date, stock_x]

        # Entry signal
        if positions[pair] == 0:
            if 1.5 < z < 2:
                # Long Y, Short X
                shares_y = int(cash[pair] / (2 * (price_y + commission)))
                shares_x = int(cash[pair] / (2 * (price_x + commission)))
                cash[pair] -= (price_y - commission) * shares_y
                cash[pair] += (price_x - commission) * shares_x
                shares[pair]['stock_y'] = shares_y
                shares[pair]['stock_x'] = -shares_x
                positions[pair] = 1

            elif -2 < z < -1.5:
                # Short Y, Long X
                shares_y = int(cash[pair] / (2 * (price_y + commission)))
                shares_x = int(cash[pair] / (2 * (price_x + commission)))
                cash[pair] += (price_y - commission) * shares_y
                cash[pair] -= (price_x - commission) * shares_x
                shares[pair]['stock_y'] = -shares_y
                shares[pair]['stock_x'] = shares_x
                positions[pair] = 1

        # Exit condition
        elif positions[pair] != 0:
            if abs(z) < 0.2 or abs(z) > 2:  # Exit: convergence or stop-loss
                shares_y = shares[pair]['stock_y']
                shares_x = shares[pair]['stock_x']
                cash[pair] += shares_y * (price_y - commission if shares_y > 0 else price_y + commission)
                cash[pair] += shares_x * (price_x - commission if shares_x > 0 else price_x + commission)
                shares[pair] = {'stock_y': 0, 'stock_x': 0}
                positions[pair] = 0

        # Update values
        if positions[pair] == 0:
            stock_value[pair] = 0
        else:
            stock_value[pair] = (shares[pair]['stock_y'] * price_y +
                                 shares[pair]['stock_x'] * price_x)

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

#Calculate Account Value
account_value_df = pd.DataFrame({
    pair: pd.DataFrame(data).set_index('Date')['sub_value']
    for pair, data in summary.items()
})
account_value_df['Total_Account_Value'] = account_value_df.sum(axis=1)

#Output everything to Excel
with pd.ExcelWriter('Trading_Summary_Output.xlsx') as writer:
    for pair, data in summary.items():
        df = pd.DataFrame(data)
        df.set_index('Date', inplace=True)
        df.to_excel(writer, sheet_name=pair)

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
plt.savefig('Total_Account_Value_Plot.png')  # Save to file
plt.show()  # Display the plot