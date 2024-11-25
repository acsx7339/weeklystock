import openpyxl
from fugle_marketdata import RestClient
from datetime import datetime, timedelta

class Update_Excel():

    def __init__(self):
        self.file_name = "stock_review.xlsx"
        self.wb = openpyxl.load_workbook(self.file_name)
        self.current_date = datetime.now().strftime("%Y-%m-%d")  # 格式化为 "YYYY-MM-DD"
        seven_days_ago = datetime.now() - timedelta(days=7)
        self.past_date = seven_days_ago.strftime("%Y-%m-%d")
        self.sheet = self.wb.create_sheet(title=self.current_date)

    def add_item(self, sdict):
        print(sdict)
        self.sheet.append(["stock", "StartPrice"])
        for stock, price in sdict.items():
            self.sheet.append([stock, price])
        self.wb.save(self.file_name)
        print(f"Excel 文件已成功创建并保存为 {self.file_name}")

    def get_price(self, stock_item):
        client = RestClient(api_key = 'NTFhZTU0YzItZDk2OC00YTI \
                    1LWE0YWQtZDhmNjVlZjA2YTY4IDFjMjkyMm \
                    JjLWYwNDktNDlmZi1hZWM2LTU4ZmYyZWEzMDE3Zg')  # 輸入您的 API key
        stock = client.stock  # Stock REST API client
        stock_dict = dict()
        for item in stock_item:
            data = stock.historical.candles(**{"symbol": item, "fields": "close"})
            print(data)
            first_item = data['data'][0]
            # first_date = first_item['date']
            first_close = first_item['close']
            print(f"stock: {item}, close price is : {first_close}")
            stock_dict[item] = first_close
        return stock_dict


if __name__ == "__main__":
    e = Update_Excel()
    item = ['2059', '2330', '2357', '2414', '2545', '3037', '3454', '6024', '6515', '6807']
    content = e.get_price(item)
    e.add_item(content)
