from GetTWStock import TWStock
from sharefolder import Finmind as fm
from candlestick import Yfinance as yf
import logging
from update_sheet import UpdateExcel as ue
from review_sheet import ReviewExcel as re

logging.basicConfig(
    level=logging.INFO,   # 設置最低的日誌級別
    format='%(levelname)s - %(message)s',  # 設置日誌格式
    handlers=[
        logging.FileHandler("Archive/result.log"),  # 將日誌寫入文件
        logging.StreamHandler()          # 將日誌輸出到控制台
    ]
)

def main():
    stocklist = TWStock().CODE()
    filter1 = []
    for item in stocklist:
        print("-----")
        logging.info(f"現在測試的是 {item}")
        try:
            stock_content = yf().yahoo_result(item)
            daily_content = yf().daily_status(stock_content)
            Weekly_content = yf().Weekly_status(stock_content)
            close_price = stock_content["Close"].iloc[-1].values[0]
            if (yf().upper_trend(daily_content, "5MA") and  # 只要是往上的就可以
                yf().upper_trend(daily_content, "10MA") and
                yf().upper_trend(daily_content, "60MA") and
                yf().upper_trend(Weekly_content, "5MA") and 
                yf().upper_trend(Weekly_content, "10MA") and
                yf().high_trend(daily_content) and  # 5MA要比10MA高且不能高超過1.08
                yf().high_trend(Weekly_content) and
                yf().near_price(close_price, 
                                yf().daily_status(stock_content)["5MA"].iloc[-1])  # close price is close 5MA
                # yf().volume_break(stock_content) and
                # yf().price_break(stock_content)
               ):
                if close_price >= 15:
                    logging.info(f"{item} 通過第一輪測試")
                    filter1.append(item)
        except Exception as e:
            logging.error(f"測試 {item} 時發生錯誤: {e}")
            continue

    logging.info(f"第一輪通過的有 {filter1}")
    logging.info("開始第二輪測試")
    
    filter2 = []
    for target in filter1:
        logging.info(f"現在測試的是： {target}")
        print("-----")
        try:
            if (fm().ratio_result(target) and
                fm().Revenue(target)):
                filter2.append(target)
        except Exception as e:
            logging.error(f"測試 {target} 時發生錯誤: {e}")
            continue

    logging.info(f"總共有 {len(filter2)} 個通過測試")
    logging.info(f"最終結果為： {filter2}")

    try:
        result = ue().get_price(filter2)
        ue().add_item(result)
        price = re().get_price()
        re().update_prices(price)
    except Exception as e:
        logging.error(f"處理最終結果時發生錯誤: {e}")

if __name__ == "__main__":
    main()
