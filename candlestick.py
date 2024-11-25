import talib
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
from GetTWStock import TWStock
from scipy.signal import find_peaks
import logging

logging.basicConfig(
    level=logging.INFO,   # 設置最低的日誌級別
    format='%(levelname)s - %(message)s',  # 設置日誌格式
    handlers=[
        logging.FileHandler("result.log"),  # 將日誌寫入文件
        logging.StreamHandler()          # 將日誌輸出到控制台
    ]
)


class Yfinance:

    end_date = pd.Timestamp(datetime.now()).normalize() + pd.offsets.MonthEnd(0)
    start_date = end_date - pd.DateOffset(months=60)
    logging.info(f"查詢股票的 開始日期: {start_date.date()}, 結束: {end_date.date()}")
    # def __init__(self):
        
        # print(f"start date: {self.start_date.date()}, end date: {self.end_date.date()}")

    def yahoo_result(self, stock):
        content = yf.download(f'{stock}.TW', start=self.start_date, end=self.end_date)
        return content
    
    def Weekly_status(self, content):
        Weekly_stock = content['Adj Close'].resample('W').last()
        Weekly_stock = Weekly_stock.dropna()
        Weekly_stock = Weekly_stock.astype(float)
        Weekly_stock_df = pd.DataFrame(Weekly_stock)
        Weekly_stock_df['5MA'] = talib.SMA(Weekly_stock_df['Adj Close'].values.astype(float), timeperiod=5)
        Weekly_stock_df['10MA'] = talib.SMA(Weekly_stock_df['Adj Close'].values.astype(float), timeperiod=10)
        logging.info(f"週Ｋ的ｍａ值： {Weekly_stock_df[['5MA', '10MA']].tail(2)}")
        return Weekly_stock_df[['5MA', '10MA']].tail(2)

    def daily_status(self, content):
        daily_stock = content['Adj Close']
        daily_stock = daily_stock.dropna()
        daily_stock = daily_stock.astype(float)
        daily_stock_df = pd.DataFrame(daily_stock)
        daily_stock_df['5MA'] = talib.SMA(daily_stock_df['Adj Close'].values.astype(float), timeperiod=5)
        daily_stock_df['10MA'] = talib.SMA(daily_stock_df['Adj Close'].values.astype(float)-0.2, timeperiod=10)
        daily_stock_df['60MA'] = talib.SMA(daily_stock_df['Adj Close'].values.astype(float), timeperiod=60)
        logging.info(f"日Ｋ的 ｍａ值： {daily_stock_df[['5MA', '10MA', '60MA']].tail(2)}")
        return daily_stock_df[['5MA', '10MA', '60MA']].tail(2)

    def upper_trend(self, item, ma):
        cur_date = item.index[1].strftime('%Y-%m-%d')
        prev_date = item.index[0].strftime('%Y-%m-%d')
        cur_ma = item.loc[cur_date, ma]
        prev_ma = item.loc[prev_date, ma]
        logging.info(f"current {ma} value is {cur_ma}, previous value is {prev_ma}")
        upper_result = cur_ma - prev_ma
        if upper_result > 0:
            logging.info(f"{ma} 是向上趨勢的 通過")
            return True
        else:
            logging.info(f"{ma} 是向下趨勢的 不通過")        

    def high_trend(self, item):
        cur_date = item.index[1].strftime('%Y-%m-%d')
        value_5ma = item.loc[cur_date, "5MA"]
        value_10ma = item.loc[cur_date, "10MA"]
        if value_10ma * 1.05 > value_5ma > value_10ma  :
            logging.info(f"5ma在 10ma的0.05%範圍內 通過")
            return True
        elif value_10ma * 1.05 < value_5ma:
            logging.info(f"5ma比10ma的1.05%還要大")
        elif value_10ma > value_5ma:
            logging.info("10ma比5ma還要大")
    def volume_break(self, item):
        daily_volume = item['Volume'] // 1000
        recent_volume = daily_volume.tail(21)
        high_volumes_21 = recent_volume.max()
        current_volume = recent_volume.iloc[-1]
        # print(high_volumes_21, current_volume)
        if current_volume > high_volumes_21:
            print("today volume is higher than recent 21 days volume")
            return True
        else:
            print("volume is not higher than previosu")
        
    def price_break(self, item):
        high_prices = item.tail(21)['High'].values
        peaks, _ = find_peaks(high_prices)
        if len(peaks) > 0:
            print("局部高点的日期和价格:")
            highest_peak_idx = peaks[np.argmax(high_prices[peaks])]
            print("最高局部高点的日期和价格:")
            print(item.tail(21).index[highest_peak_idx], high_prices[highest_peak_idx])
        else:
            print("没有找到局部高点")
        cur_close = item['Close'][-1]
        # print(cur_close)
        if cur_close > high_prices[highest_peak_idx]:
            return True 
        else:
            print("price is not higher than previous")

    def near_price(self, price, ma):
        if ma * 1.05 >= price >= ma * 1.02:
            logging.info("收盤價在5ma附近")
            return True
        else:
            logging.info("收盤價不在5ma附近")


        
# if __name__ == "__main__":
    # stock = Yfinance()
    # content = stock.yahoo_result("2102")
    # print(stock.daily_status(content)["5MA"].iloc[-1])
    # print(content['Close'].iloc[-1])
    # stock.price_break(content)
    # print(content)
    # stock.volume(content)
    # item = stock.Weekly_status(content)
    # stock.upper_trend(item, "5MA")
    # stock.high_trend(item)
