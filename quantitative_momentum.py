import numpy as np
import pandas as pd

import math
import requests
import xlsxwriter

from scipy.stats import percentileofscore as percentile
from secrets import IEX_CLOUD_API_TOKEN
from utils import chunks

stocks = pd.read_csv('sp_500_stocks.csv')

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for group in symbol_groups:
    symbol_strings.append(','.join(group))

high_quality_momentum_cols = ['Ticker', 'Price', 'Number of Shares to Buy', 'One-Year Price Return', 'One-Year Return Percentile', 'Six-Month Price Return', 'Six-Month Return Percentile', 'Three-Month Price Return', 'Three-Month Return Percentile', 'One-Month Price Return', 'One-Month Return Percentile', 'High Quality Momentum Score']
hqm_dataframe = pd.DataFrame(columns=high_quality_momentum_cols)
for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}&types=stats,price'
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = hqm_dataframe.append(pd.Series([symbol, data[symbol]['price'], 'N/A', data[symbol]['stats']['year1ChangePercent'], 'N/A', data[symbol]['stats']['month6ChangePercent'], 'N/A', data[symbol]['stats']['month3ChangePercent'], 'N/A', data[symbol]['stats']['month1ChangePercent'], 'N/A', 'N/A'], index=high_quality_momentum_cols), ignore_index=True)

time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']
for row in hqm_dataframe.index:
    for time_period in time_periods:
        return_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_dataframe.loc[row, percentile_col] = percentile(hqm_dataframe[return_col], hqm_dataframe.loc[row, return_col])

for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'High Quality Momentum Score'] = .4 * momentum_percentiles[0] + .3 * momentum_percentiles[1] + .2 * momentum_percentiles[2] + .1 * momentum_percentiles[3]

hqm_dataframe.sort_values('High Quality Momentum Score', ascending=False, inplace=True)
hqm_dataframe = hqm_dataframe[:50]
hqm_dataframe.reset_index(inplace=True, drop=True)

val = None
while type(val) is not float:
    portfolio_size = input('Enter the cash vaue of your portfolio: ')
    try:
        val = float(portfolio_size)
    except ValueError:
        print('That\'s not a number!')

for i in range(len(hqm_dataframe.index)):
    position_size = val / (len(hqm_dataframe.index) - i)
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / hqm_dataframe.loc[i, 'Price'])
    val -= hqm_dataframe.loc[i, 'Number of Shares to Buy'] * hqm_dataframe.loc[i, 'Price']

writer = pd.ExcelWriter('recommended_trades_quantitative_momentum.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, 'Recommended Trades', index=False)

bg_color = '#a5d8dd'
font_color = '#20283e'
string_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1})
dollar_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'$0.00'})
integer_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'0'})
percent_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'0.00%'})
float_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'0.00'})

column_formats = {'A': ['Ticker', string_format], 'B': ['Price', dollar_format], 'C': ['Number of Shares to Buy', integer_format], 'D': ['One-Year Price Return', percent_format], 'E': ['One-Year Return Percentile', float_format], 'F': ['Six-Month Price Return', percent_format], 'G': ['Six-Month Return Percentile', float_format], 'H': ['Three-Month Price Return', percent_format], 'I': ['Three-Month Return Percentile', float_format], 'J': ['One-Month Price Return', percent_format], 'K': ['One-Month Return Percentile', float_format], 'L': ['High Quality Momentum Score', float_format]}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)
writer.save()
