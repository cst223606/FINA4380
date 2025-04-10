#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd
import numpy as np
import itertools
from scipy import odr
from statsmodels.tsa.stattools import adfuller


# Load Excel file
df = pd.read_excel(r'\STOCKS.xlsx', parse_dates=['Date'], index_col='Date')


# Check
print(df.head())


# Calculate log-price ratios
log_price_ratios = np.log(df / df.iloc[0])

# Define TLS regression without intercept
def tls_no_intercept(x, y):
    def linear_func(beta, x):
        return beta * x
    linear_model = odr.Model(linear_func)
    data = odr.RealData(x, y)
    odr_reg = odr.ODR(data, linear_model, beta0=[1.0])
    odr_out = odr_reg.run()
    beta = odr_out.beta[0]
    residuals = y - beta * x
    return beta, residuals

# Define ADF test function with automatic lag selection
def adf_test_with_lag_selection(residuals, ic='AIC'):
    result = adfuller(residuals, maxlag=None, regression='c', autolag=ic)
    return result[0], result[1], result[2]  # adf_stat, p_value, used_lag

# Perform pairwise analysis
stocks = log_price_ratios.columns
results = []

for stock_x, stock_y in itertools.combinations(stocks, 2):
    x = log_price_ratios[stock_x].values
    y = log_price_ratios[stock_y].values
    
    # TLS regression
    beta, residuals = tls_no_intercept(x, y)
    
    # ADF test on residuals with lag selection (AIC)
    adf_stat, p_value, used_lag = adf_test_with_lag_selection(residuals, ic='AIC')  
    
    # Store results
    results.append({
        'Stock_X': stock_x,
        'Stock_Y': stock_y,
        'Beta': beta,
        'ADF_Statistic': adf_stat,
        'ADF_pvalue': p_value,
        'ADF_used_lag': used_lag
    })

# Convert results to DataFrame
results_df = pd.DataFrame(results)

# Identify cointegrated pairs (p-value < 0.05)
cointegrated_pairs = results_df[results_df['ADF_pvalue'] < 0.05].sort_values('ADF_pvalue')

print(cointegrated_pairs.head())


# In[ ]:




