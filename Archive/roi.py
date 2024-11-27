from fugle_marketdata import RestClient

client = RestClient(api_key = 'NTFhZTU0YzItZDk2OC00YTI \
                    1LWE0YWQtZDhmNjVlZjA2YTY4IDFjMjkyMm \
                    JjLWYwNDktNDlmZi1hZWM2LTU4ZmYyZWEzMDE3Zg')  # 輸入您的 API key
stock = client.stock  # Stock REST API client

stock_item = ['2402', '2419', '8028']
for item in stock_item:
    data = stock.historical.candles(**{"symbol": item, "fields": "close"})
    # print(data)
    first_item = data['data'][0]
    first_date = first_item['date']
    first_close = first_item['close']
    print(f"stock: {item}, close price is : {first_close}")