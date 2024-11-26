import openpyxl
from fugle_marketdata import RestClient
from datetime import datetime, timedelta
import twstock

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
        stock_dict = dict()
        for item in stock_item:
            stock = twstock.Stock(item)
            current_price = stock.price[-1]
            stock_dict[item] = current_price
        return stock_dict


if __name__ == "__main__":
    e = Update_Excel()
    item = ['2059', '2330', '2357', '2414', '2545', '3037', '3454', '6024', '6515', '6807']
    content = e.get_price(item)
    e.add_item(content)
