import numpy as np
import pandas as pd
import yfinance as yf
from scipy import stats
import math

#Read the constituent stock of S&P 500
stocks = pd.read_csv('constituents.csv')

#Making the first call yFinance
AAPL = yf.Ticker('AAPL').info

price = AAPL['regularMarketPreviousClose']
PE_ratio = AAPL['trailingPE']

#Define the column of the dataframe
my_columns = ['Ticker', 'Price', 'Price-to-Earning Ratio', 'Number of Shares to Buy']

#Adding temporary array too hold data
L = []
x = 1

#Loop thorough all S&P 500 Constituent
for stock in stocks['Symbol']:

    #Get data from Yahoo Finance API
    data = yf.Ticker(stock).info

    #Determine if the P/E Ratio is existing or not
    try:
        print(f'{x} {data["trailingPE"]}')
        info = pd.Series([stock, data['regularMarketPreviousClose'], data['trailingPE'], 'N/A'], index=my_columns)
    except KeyError:
        info = pd.Series([stock, data['regularMarketPreviousClose'],100000 , 'N/A'], index = my_columns)
        print(info)
    L.append(info)
    x = x + 1

#Add the temporary data holder to the dataframe
final_dataframe = pd.DataFrame(L, columns = my_columns)

#Ranking the dataframe by Ascending Price to Earning Ratio and keep only the top 50
final_dataframe = final_dataframe.sort_values('Price-to-Earning Ratio', ascending = True)
final_dataframe = final_dataframe.reset_index(drop = True)
final_dataframe = final_dataframe.loc[:50]

#getting portfolio size
portfolio_size = input("Enter the value of your portfolio")

#Try eliminating out error
try:
    value = float(portfolio_size)
except ValueError:
    print(f"This is not a number\nTry Again Please")
    portfolio_size = input("Enter the value of your portfolio")
    value = float(portfolio_size)

#Calculate position size for each individual stocks
position_size = value/len(final_dataframe.index)

#Calculate the number of share to buy for each stocks
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Price'])

