import numpy as np
import pandas as pd
import xlsxwriter
import yfinance as yf
from scipy import stats
import math

stocks = pd.read_csv('constituents.csv')

#Making the first call yFinance
AAPL = yf.Ticker('AAPL').info

price = AAPL['regularMarketPreviousClose']
PE_ratio = AAPL['trailingPE']

my_columns = ['Ticker', 'Price', 'Price-to-Earning Ratio', 'Number of Shares to Buy']
L = []
x = 1
for stock in stocks['Symbol']:
    data = yf.Ticker(stock).info

    try:
        print(f'{x} {data["trailingPE"]}')
        info = pd.Series([stock, data['regularMarketPreviousClose'], data['trailingPE'], 'N/A'], index=my_columns)
    except KeyError:
        info = pd.Series([stock, data['regularMarketPreviousClose'],100000 , 'N/A'], index = my_columns)
        print(info)
    L.append(info)
    x = x + 1

final_dataframe = pd.DataFrame(L, columns = my_columns)

final_dataframe = final_dataframe.sort_values('Price-to-Earning Ratio', ascending = True)
final_dataframe = final_dataframe.reset_index(drop = True)
final_dataframe = final_dataframe.drop([50, len(final_dataframe)-1])
print(final_dataframe)


