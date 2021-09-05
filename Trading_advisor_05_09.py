# -*- coding: utf-8 -*-
"""
Created on Sun Sep  5 19:24:10 2021

@author: andre
"""


import trendln
# this will serve as an example for security or index closing prices, or low and high prices
import yfinance as yf 
import pandas as pd
from openpyxl import load_workbook

# array with all available stocks
stocks_pool = []

xl = pd.ExcelFile('C:\Trading\Акции доступные в Тинькофф.xlsx')
stocks_in_file = xl.parse('Тикеры')

i = 3
for i in stocks_in_file.index:    
    stocks_pool.append(stocks_in_file.loc[i, 'Ticker']) # indexes of the list match the indexes of the Dataframe


tick = yf.Ticker('TSLA') 
hist = tick.history(period="max", rounding=True)
# =============================================================================
# h = hist[-1000:].Close
# mins, maxs = trendln.calc_support_resistance(h, accuracy=8)
# fig = trendln.plot_support_resistance(hist[-1000:].Close, accuracy=8)
# =============================================================================

