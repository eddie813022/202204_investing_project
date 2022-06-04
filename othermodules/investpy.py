import investpy

# country_list = investpy.stocks.get_stock_countries() # 搜尋國家清單->數量91核對沒錯
# result = investpy.stocks.get_stocks("taiwan") # 根據國家檢索股票數據



# result = investpy.stocks.get_stocks_list("taiwan") # 根據國家返回股票代碼list

# investpy.stocks.get_stocks_overview # 根據國家返回及時數據總覽
# investpy.stocks.get_stocks_dict # 根據國家檢索股票數據，與 investpy.stocks.get_stocks("taiwan") 相同
# df2 = investpy.stocks.get_stocks_dict("taiwan")
# for i in df2:
    # if i["symbol"] == "2317":
    #     print(i["name"])
# {'country': 'taiwan', 'name': 'MetaTech AP', 'full_name': 'MetaTech AP', 'isin': 'TW0003224002', 'currency': 'TWD', 'symbol': '3224'}
# 歷史紀錄
# df = investpy.get_stock_historical_data(stock='2317',
#                                         country='taiwan',
#                                         from_date='01/01/2017',
#                                         to_date='01/01/2022',
#                                         interval="Monthly")

# df = investpy.stocks.get_stock_dividends(stock="2317",country='taiwan') # 股利(台灣異常)
# df = investpy.stocks.get_stock_financial_summary("2317","taiwan",summary_type="balance_sheet",period="annual") # 財務摘要
# annual quarterly
# df = investpy.stocks.get_stock_recent_data("2317","taiwan") # 近期歷史數據(昨日~一個月)
# df = investpy.stocks.get_stock_information("2317","taiwan") # 當日綜觀


# df = investpy.technical.moving_averages(name='2317', country='taiwan', product_type='stock', interval='daily')
# df = investpy.technical.pivot_points(name='2317', country='taiwan', product_type='stock', interval='daily')
# df = investpy.technical.technical_indicators(name='2317', country='taiwan', product_type='stock', interval='daily')

# df = investpy.search.search_quotes(text='2317', products=['stocks'],
#                                        countries=['taiwan'], n_results=1)
# print(df.retrieve_technical_indicators())
# df.to_excel("history2.xlsx")










