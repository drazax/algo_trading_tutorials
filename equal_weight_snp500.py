import numpy as np
import pandas as pd
import math
import requests
import xlsxwriter

stocks = pd.read_csv('sp_500_stocks.csv')

from secrets import IEX_CLOUD_API_TOKEN

def chunks(l, n):
    # yield successive n-sized chunks from list l
    for i in range(0, len(l), n):
        yield l[i:i + n]

my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for group in symbol_groups:
    symbol_strings.append(','.join(group))

for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(pd.Series([symbol, data[symbol]['quote']['latestPrice'], data[symbol]['quote']['marketCap'], 'N/A'], index=my_columns), ignore_index=True)

val = None
while type(val) is not float:
    portfolio_size = input('Enter the cash value of your portfolio: ')
    try:
        val = float(portfolio_size)
    except ValueError:
        print('That\'s not a number!')

position_size = val / len(final_dataframe.index)
for i in range(len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / final_dataframe.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index=False)

bg_color = '#a5d8dd'
font_color = '#20283e'
string_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1})
dollar_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'$0.00'})
integer_format = writer.book.add_format({'font_color':font_color, 'bg_color':bg_color, 'border':1, 'num_format':'0'})

column_formats = {'A': ['Ticker', string_format], 'B': ['Stock Price', dollar_format], 'C': ['Market Capitalization', dollar_format], 'D': ['Number of Shares to Buy', integer_format]}
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)
writer.save()
