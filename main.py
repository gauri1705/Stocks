from FinData import FinancialData

#tickers_list = ['^BSESN', '^NSEI','^NSEMDCP50','^NSESMCP50']
f_data = FinancialData()
print(f_data.tickers)
f_data.load_all_tickers()
f_data.save_data('findata01.xlsx')

