# 載入 yfinance 套件
# 安裝方式：pip install yfinance
# 查看版本：pip show yfinance
# 升級套件：pip install --upgrade yfinance
# 移除方式：pip uninstall yfinance
import yfinance as yf

# 載入 FinMind 套件
# 安裝方式：pip install FinMind
# https://finmindtrade.com/analysis/#/account/user
# https://finmindtrade.com/analysis/#/python_gym/finmind_tech
# https://github.com/FinMind/FinMind?tab=readme-ov-file
from FinMind.data import DataLoader

# 載入 twstock 套件
# 安裝方式：pip install twstock
# https://github.com/mlouielu/twstock
# https://twstock.readthedocs.io/zh-tw/latest/index.html
import twstock

# 用來檢查檔案是否存在
import os

# 載入環境變數工具
# 安裝方式：pip install python-dotenv
from dotenv import load_dotenv
load_dotenv() # 先導入再調用

# 初始化並認證 FinMind DataLoader
dl = DataLoader()
dl.login_by_token(api_token=os.getenv('FINMIND_API_TOKEN'))  # 一次性認證

# 載入 繪圖 套件
# 安裝方式：pip install matplotlib
import matplotlib.pyplot as plt

# 偵測系統平台
import platform

# 日記記錄
import logging

# 時間格式
from datetime import datetime, timedelta, time

# 安裝方式：pip install pandas
# 安裝方式：python.exe -m pip install --upgrade pip
import pandas as pd
pd.set_option('display.max_columns', None)  # 在檔案頂部添加

# 安裝方式：pip install requests
import requests

# 安裝方式：pip install beautifulsoup4
from bs4 import BeautifulSoup

from utils import setup_logger, setup_check_realtime_format, setup_folders, setup_chinese_font  # 引入共用函數

import tkinter as tk
from tkinter import messagebox

import inspect
#print(f"函式的名字是：{inspect.currentframe().f_code.co_name}")

import math

# 安裝方式：pip install selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time as time_module  # 用於計時功能

# 富果台股即時行情 API
# 安裝方式：pip install fugle-marketdata
from fugle_marketdata import WebSocketClient, RestClient

# ChatGPT API 設定
API_KEY = "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"  # 請替換成你的 API Key
API_ENDPOINT = "https://api.openai.com/v1/chat/completions"

def process_stock_data(ticker='2330', start_date='', end_date=''):
    """
    下載並處理股票數據，繪製圖表並儲存 CSV。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'
        start_date (str): 開始日期，例如 '2025-02-17'
        end_date (str): 結束日期，例如 '2025-02-26'
        參考網址：https://hk.finance.yahoo.com/quote/0052.TW/history/
        參考網址：https://colab.research.google.com/gist/pythonviz/6ae108a75c5ab44efbcd7d6a64848017/yahoo-finance-in-python.ipynb
        參考網址：https://yfinance-python.org/
        參考網址：https://github.com/ranaroussi/yfinance/blob/0713d9386769b168926d3959efd8310b56a33096/yfinance/utils.py#L445-L462
    """

    # 今天日期
    today = datetime.now().strftime('%Y-%m-%d')
    
    # 如果開始時間未定義，預設為今天往前減 30 天
    if not start_date:
        start_date = today
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        start_date = (start_date + timedelta(days=-30)).strftime('%Y-%m-%d')
    
    # 如果結束時間未定義，預設為今天往後加 1 天，否則對輸入的 end_date 往後加 1 天
    if not end_date:
        end_date_dt = datetime.strptime(today, '%Y-%m-%d')
    else:
        end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')

    end_date = (end_date_dt + timedelta(days=1)).strftime('%Y-%m-%d')

    # 檔名識別
    identify_name = 'yfinance_daily'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯日期行為
    setup_check_realtime_format(logger, start_date, end_date, '>=')

    # 簡化日期格式
    start_short = start_date.replace('-', '')
    new_end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')
    new_end_date = (new_end_date_dt + timedelta(days=-1)).strftime('%Y-%m-%d')
    end_short = new_end_date.replace('-', '')

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載股票數據並清理
    try:
        # data = yf.download(f"{ticker}.TW", start=start_date, end=end_date)
        # 判斷 ticker 是否為數字，若是則加上 .TW，否則直接使用 ticker
        stock_suffix = ''
        if ticker.isdigit():
            stock_suffix = '.TW'
        stock = yf.Ticker(f"{ticker}" + stock_suffix)
        data = stock.history(start=start_date, end=end_date, auto_adjust=False)
        data['Change'] = data['Close'].diff().round(2)
        
        # 移除時區資訊並將日期轉為普通欄位
        data = data.reset_index()
        data['Date'] = data['Date'].dt.date  # 移除時區，僅保留日期
        
        # 選擇並重新命名欄位，保持與原格式一致
        # data = data[['Date', 'Close', 'High', 'Low', 'Open', 'Volume']]
        data['Date'] = pd.to_datetime(data['Date'])  # 確保日期格式
    except Exception as e:
        logger.error(f"下載 {ticker} 數據異常：{e}")
        return None  # 若發生連線異常，返回 None 並結束函數

    # 檢查並顯示數據
    if data.empty:
        logger.warning(f"沒有找到 {ticker} 的數據，可能是日期範圍無效或未來日期")
        return None  # 若無數據，返回 None 並結束函數

    # 將價格欄位四捨五入到小數點後 2 位
    price_columns = ['Open', 'High', 'Low', 'Close', 'Adj Close']
    data[price_columns] = data[price_columns].round(2)

    # 互換 Close 和 Open 欄位（註解中，保持原樣）
    # data[['Close', 'Open']] = data[['Open', 'Close']]

    # 按日期遞減排序（註解中，保持原樣）
    # data = data.sort_values('Date', ascending=False)

    # 格式化 Volume 欄位，加入千位分隔符
    data['Volume'] = data['Volume'].apply(lambda x: f"{int(x):,}")

    # 繪製折線圖，明確指定 X 軸數據
    plt.plot(data['Date'], data['Close'], marker='o', linestyle='-', color='blue', label='股價')
    # plt.plot(data['Date'], data['Adj Close'], marker='o', linestyle='-', color='red', label='調整後的收盤價')
    plt.title(f'{ticker} 股價趨勢')
    plt.grid(True)  # 添加網格線
    plt.xlabel('日期')
    plt.ylabel('股價')
    # plt.legend(title='價格類型')
    plt.xticks(data['Date'], rotation=90)

    # 在每個圓點旁顯示值 - Close
    for i, (date, value, change) in enumerate(zip(data['Date'], data['Close'], data['Change'])):
        label = f'{value:.2f}（{change:+.2f}）'  # 顯示 "收盤價（+漲跌）"，+號顯示正負
        va = 'bottom' if change >= 0 else 'top'  # 正數在上方，負數在下方
        plt.text(date, value, label, ha='left', va=va, fontsize=8, 
                 color='red' if change >= 0 else 'green')  # 正數紅色，負數綠色

    # 在每個圓點旁顯示值 - Adj Close
    # for i, (date, value) in enumerate(zip(data['Date'], data['Adj Close'])):
        # plt.text(date, value, f'{value:.2f}', ha='left', va='top', fontsize=8, color='red')

    """
    # 只標註最高和最低價
    max_idx = data['Close'].idxmax()
    min_idx = data['Close'].idxmin()
    plt.text(max_idx, data['Close'][max_idx], f'{data["Close"][max_idx]:.2f}', ha='left', va='bottom', fontsize=8)
    plt.text(min_idx, data['Close'][min_idx], f'{data["Close"][min_idx]:.2f}', ha='left', va='top', fontsize=8)
    """

    # 更改欄位名稱為中文
    data = data.rename(columns={'Date': '日期', 'Open': '開盤價', 'High': '最高價', 'Low': '最低價', 'Close': '收盤價', 'Adj Close': '調整後的收盤價', 'Volume': '成交量', 'Dividends': '股利', 'Stock Splits': '股票分割', 'Change': '漲跌'})

    # 儲存 CSV 檔案
    csv_file = file_name + '.csv'
    if os.path.exists(csv_file):
        logger.warning(f"csv警告：{csv_file} 已存在，將被覆蓋")
    data.to_csv(csv_file, index=False)  # 添加 index=False

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    logger.debug(f"{ticker}_{identify_name}_{start_short}_{end_short} 資料數據如下")  # 細節用 DEBUG
    logger.info(f"\n{data.to_string(index=False)}")

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data  # 返回處理後的 DataFrame，供後續使用

def night_trading(ticker='WCDFP&'):
    """
    台積電期貨2504
    網址：https://www.wantgoo.com/futures/wcdfj5

    台積電期貨
    網址：https://www.wantgoo.com/futures/wcdf&

    台積電期貨盤後
    網址：https://www.wantgoo.com/futures/wcdfp&

    台積電期貨合併
    網址：https://www.wantgoo.com/futures/wcdfm&

    台指期2504
    網址：https://www.wantgoo.com/futures/wtxj5

    台指期
    網址：https://www.wantgoo.com/futures/wtx&

    台指期盤後
    網址：https://www.wantgoo.com/futures/wtxp&

    台指期合併
    網址：https://www.wantgoo.com/futures/wtxm&

    https://www.taifex.com.tw/cht/3/dlFutDailyMarketView
    https://www.taifex.com.tw/cht/3/futDailyMarketReport
    """

    # 檔名識別
    identify_name = ticker.replace('&', '')

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}'

    # 關聯字型行為
    setup_chinese_font(logger)

    api_data_wantgoo = wantgoo_all_quote_info(ticker)
    
    # 提取時間戳
    last_timestamp = int(api_data_wantgoo.get('time', 0))
    # 將時間戳轉為 "年-月-日" 格式
    last_date_str = (datetime.fromtimestamp(last_timestamp / 1000) + timedelta(days=-1)).strftime('%Y-%m-%d')
    
    # 構建 DataFrame
    close = float(api_data_wantgoo.get('close', 0))
    previousClose = float(api_data_wantgoo.get('previousClose', 0))
    spread = close - previousClose
    history_data = pd.DataFrame({
        # 'time': [last_timestamp],  # 原始時間戳
        'Date': [last_date_str],  # 新增：轉換後的日期
        'open': [float(api_data_wantgoo.get('open', 0))],
        'high': [float(api_data_wantgoo.get('high', 0))],
        'low': [float(api_data_wantgoo.get('low', 0))],
        'close': [close],
        # 'volume': [int(api_data_wantgoo.get('volume', 0))],
        'previousClose': [previousClose],
        # 'previousVolume': [int(api_data_wantgoo.get('previousVolume', 0))],
        'spread': [spread]
    })

    # 歷史數據檔案路徑
    history_file = f'output/{ticker}/{ticker}.csv'
    
    if not os.path.exists(history_file):
        # 首次寫入，帶表頭
        history_data.to_csv(history_file, index=False)
        logger.info(f"首次創建檔案並寫入數據：{last_date_str}")
    else:
        # 讀取現有數據
        existing_data = pd.read_csv(history_file)
        # 檢查是否已有相同 Date
        if last_date_str not in existing_data['Date'].values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, mode='a', header=False, index=False)  # 追加，不帶表頭
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")


    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])  # 轉為 datetime

    # 繪製折線圖，明確指定 X 軸數據
    plt.plot(history_data_load['Date'], history_data_load['close'], marker='o', linestyle='-', color='blue', label='股價')
    plt.title(f'{ticker} 股價趨勢')
    plt.grid(True)  # 添加網格線
    plt.xlabel('日期')
    plt.legend(title='價格類型')
    plt.xticks(history_data_load['Date'], rotation=90)

    # 在每個圓點旁顯示值 - close
    for i, (date, close, spread) in enumerate(zip(history_data_load['Date'], history_data_load['close'], history_data_load['spread'])):
        label = f'{close:.2f} ({spread:+.2f})'  # 顯示 "收盤價 (+漲跌)"，+號顯示正負
        va = 'bottom' if spread >= 0 else 'top'  # 正數在上方，負數在下方
        plt.text(date, close, label, ha='left', va=va, fontsize=8, 
                 color='red' if spread >= 0 else 'green')  # 正數紅色，負數綠色

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

def exchange_rate(currency='USD', logger=None):
    """
    從台灣銀行網站抓取指定貨幣（預設 USD）的歷史匯率數據，並儲存為 CSV。
    
    Args:
        currency (str): 目標貨幣代碼，預設為 'USD'。
        logger (logging.Logger, optional): 用於記錄日誌的物件。
    
    Returns:
        dict: 包含最新一筆匯率數據的字典，若失敗則返回 None。
        參考網址：https://rate.bot.com.tw/xrt/quote/ltm/USD
        參考網址：https://rate.bot.com.tw/xrt/all/2025-04-01
    """
    url = f"https://rate.bot.com.tw/xrt/quote/ltm/{currency}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        # 發送請求並確保成功
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 解析 HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 找到目標表格
        table = soup.find('table', class_='table table-striped table-bordered table-condensed table-hover')
        if not table:
            if logger:
                logger.error("無法找到目標匯率表格")
            return None
        
        # 提取表格中的所有行
        rows = table.find('tbody').find_all('tr')
        if not rows:
            if logger:
                logger.error("表格中無資料行")
            return None
        
        # 儲存數據的列表
        data_list = []
        
        # 遍歷每一行提取數據
        for row in rows:
            tds = row.find_all('td')
            if len(tds) < 6:
                if logger:
                    logger.error("資料行欄位數不足")
                continue
            
            # 提取日期並轉換格式（從 "2025/04/01" 轉為 "2025-04-01"）
            raw_date = tds[0].text.strip()
            formatted_date = raw_date.replace('/', '-')
            
            # 提取即期匯率的本行買入與本行賣出
            spot_buy = tds[4].text.strip()  # 本行買入（即期）
            spot_sell = tds[5].text.strip() # 本行賣出（即期）
            
            # 檢查數據是否為有效數字
            if spot_buy == '-' or spot_sell == '-':
                if logger:
                    logger.debug(f"日期 {formatted_date} 的匯率數據無效，跳過")
                continue
            
            # 轉換為浮點數
            data = {
                '日期': formatted_date,
                '本行買入（外資進場買股，數值越高越利多）': float(spot_buy),
                '本行賣出（外資撤資賣股，數值越高越利空）': float(spot_sell)
            }
            data_list.append(data)
        
        if not data_list:
            if logger:
                logger.error("無有效匯率數據可提取")
            return None
        
        # 轉為 DataFrame
        df = pd.DataFrame(data_list)
        df = df.sort_values(by='日期', ascending=True)  # 按日期從早到晚排序
        
        # 寫入 CSV（直接覆蓋）
        output_dir = f'output/exchange_rate/{currency}'
        os.makedirs(output_dir, exist_ok=True)  # 確保目錄存在
        history_file = f'{output_dir}/{currency}_exchange_rate.csv'
        df.to_csv(history_file, index=False, encoding='utf-8')
        
        if logger:
            logger.info(f"成功抓取並儲存 {currency} 匯率數據至 {history_file}，共 {len(df)} 筆")
        
        # 返回最新一筆數據（假設第一行為最新）
        latest_data = data_list[0]
        return latest_data
    
    except requests.RequestException as e:
        if logger:
            logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        if logger:
            logger.error(f"數據解析錯誤：{str(e)}")
        return None

def number_of_empty_orders(queryDate=''):
    """
    從台灣期貨交易所網站抓取外資每日空單口數，並儲存為 CSV。
    
    Args:
        queryDate (str): 查詢日期，格式為 '年/月/日' (如 '2025/04/01')，預設為空值，若空則使用當前日期。
        logger (logging.Logger, optional): 用於記錄日誌的物件。
    
    Returns:
        dict: 包含指定日期的外資空單數據，若失敗則返回 None
        參考網址：https://www.wantgoo.com/futures/institutional-investors/net-open-interest
        參考網址：https://www.taifex.com.tw/cht/3/futContractsDate
        參考網址：https://stock.wearn.com/taifexphoto.asp
    """
    # 處理日期：若未指定，使用當前日期
    if not queryDate:
        queryDate = datetime.now().strftime('%Y/%m/%d')  # 格式：2025/04/01
        start_date = datetime.strptime(queryDate, '%Y/%m/%d')
        start_date = (start_date + timedelta(days=-1)).strftime('%Y/%m/%d')
    formatted_date = start_date.replace('/', '-')  # 轉為 2025-04-01 供 CSV 使用
    
    # 關聯log行為
    logger = setup_logger('market', 'number_of_empty_orders')
    
    url = "https://www.taifex.com.tw/cht/3/futContractsDate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    payload = {
        'queryType': '1',
        'doQuery': '1',
        'queryDate': start_date,
        'commodityId': 'TXF'  # 台指期貨
    }
    
    try:
        # 發送 POST 請求並確保成功
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        response.raise_for_status()
        
        # 解析 HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 找到目標表格
        table = soup.find('table', class_='table_f table-sticky-3 w-1000')
        if not table:
            if logger:
                logger.error("無法找到 class='table_f table-sticky-3 w-1000' 的表格")
            return None
        
        # 找到 TBODY
        tbody = table.find('tbody')
        if not tbody:
            if logger:
                logger.error("表格中無 TBODY")
            return None
        
        # 提取所有 TR，並找到第三個 TR（外資數據）
        trs = tbody.find_all('tr')
        if len(trs) < 3:
            if logger:
                logger.error("TR 數量不足，無法找到外資數據行")
            return None
        
        target_tr = trs[2]  # 第三個 TR（索引從 0 開始）
        tds = target_tr.find_all('td')
        if len(tds) < 12:
            if logger:
                logger.error("目標 TR 的 TD 數量不足")
            return None
        
        # 提取倒數第二個 TD 中的 span 值（外資空單淨額）
        target_td = tds[-2]  # 倒數第二個 TD
        span = target_td.find('span', class_='blue')
        if not span:
            if logger:
                logger.error("目標 TD 中無 span.blue 元素")
            return None
        
        empty_orders = span.text.strip().replace(',', '')  # 去除空白與逗號，例如 "-29,385" -> "-29385"
        
        # 構建數據
        data = {
            '日期': formatted_date,
            '外資空單口數': int(empty_orders)  # 轉為整數
        }
        
        
        # 寫入 CSV
        output_dir = 'output/futures'
        os.makedirs(output_dir, exist_ok=True)  # 確保目錄存在
        history_file = f'{output_dir}/foreign_empty_orders.csv'
        
        # 檢查是否已有檔案並比對日期
        if os.path.exists(history_file):
            existing_data = pd.read_csv(history_file)
            if formatted_date in existing_data['日期'].astype(str).values:
                if logger:
                    logger.info(f"日期 {formatted_date} 已存在，跳過寫入")
                return data  # 返回數據但不寫入
            else:
                # 追加新數據
                history_data = pd.DataFrame([data])
                history_data.to_csv(history_file, mode='a', header=False, index=False)
                if logger:
                    logger.info(f"追加新數據：{formatted_date}, 外資空單口數: {empty_orders}")
        else:
            # 首次寫入
            history_data = pd.DataFrame([data])
            history_data.to_csv(history_file, index=False)
            if logger:
                logger.info(f"首次寫入數據：{formatted_date}, 外資空單口數: {empty_orders}")
        
        return data
    
    except requests.RequestException as e:
        if logger:
            logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        if logger:
            logger.error(f"數據解析錯誤：{str(e)}")
        return None

def taiwan_stock_daily(ticker='2330', start_date='', end_date=''):
    """
    下載並處理股票數據，繪製圖表並儲存 CSV。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'
        start_date (str): 開始日期，例如 '2025-02-17'
        end_date (str): 結束日期，例如 '2025-02-26'
        參考網址：https://www.twse.com.tw/zh/trading/historical/stock-day.html
        參考網址：https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20250401&stockNo=2330&response=json
        參考網址：https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20250313&stockNo=2330&response=html
        API使用方式：https://finmind.github.io/tutor/TaiwanMarket/Technical/#taiwanstockprice
    """
        
    # 今天日期
    today = datetime.now().strftime('%Y-%m-%d')
    
    # 如果開始時間未定義，預設為今天往前減 30 天
    if not start_date:
        start_date = today
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        start_date = (start_date + timedelta(days=-30)).strftime('%Y-%m-%d')
    
    # 如果結束時間未定義，預設為今天，否則為原來的 end_date
    if not end_date:
        end_date_dt = datetime.strptime(today, '%Y-%m-%d')
    else:
        end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')

    end_date = (end_date_dt + timedelta(days=0)).strftime('%Y-%m-%d')

    # 檔名識別
    identify_name = 'finmind_daily'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯日期行為
    setup_check_realtime_format(logger, start_date, end_date, '>')

    # 簡化日期格式
    start_short = start_date.replace('-', '')
    end_short = end_date.replace('-', '')

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載股票數據並清理（使用全局 dl）
    try:
        data = dl.taiwan_stock_daily(stock_id=ticker, start_date=start_date, end_date=end_date)
        data['date'] = pd.to_datetime(data['date'])
        data = data.reset_index(drop=True)  # 移除數字索引
    except Exception as e:
        logger.error(f"下載 {ticker} 歷史股價數據異常：{e}")
        return None  # 若發生連線異常，返回 None 並結束函數

    # 檢查並顯示數據
    if data.empty:
        logger.warning(f"沒有找到 {ticker} 的歷史股價數據，可能是日期範圍無效或未來日期")
        return None  # 若無數據，返回 None 並結束函數

    # 格式化 Trading_Volume 和 Trading_money 欄位，加入千位分隔符
    data['Trading_Volume'] = data['Trading_Volume'].apply(lambda x: f"{x:,}")
    data['Trading_money'] = data['Trading_money'].apply(lambda x: f"{x:,}")

    # 繪製折線圖，明確指定 X 軸數據
    plt.plot(data['date'], data['close'], marker='o', linestyle='-', color='blue')
    plt.title(f'{ticker} 股價趨勢')
    plt.grid(True)  # 添加網格線
    plt.xlabel('日期')
    plt.ylabel('股價')
    plt.xticks(data['date'], rotation=90)

    # 在每個圓點旁顯示收盤價和漲跌
    for i, (date, close, spread) in enumerate(zip(data['date'], data['close'], data['spread'])):
        label = f'{close:.2f} ({spread:+.2f})'  # 顯示 "收盤價 (+漲跌)"，+號顯示正負
        va = 'bottom' if spread >= 0 else 'top'  # 正數在上方，負數在下方
        plt.text(date, close, label, ha='left', va=va, fontsize=8, 
                 color='red' if spread >= 0 else 'green')  # 正數紅色，負數綠色

    """
    # 只標註最高和最低價
    max_idx = data['close'].idxmax()
    min_idx = data['close'].idxmin()
    plt.text(max_idx, data['close'][max_idx], f'{data["close"][max_idx]:.2f}', ha='left', va='bottom', fontsize=8)
    plt.text(min_idx, data['close'][min_idx], f'{data["close"][min_idx]:.2f}', ha='left', va='top', fontsize=8)
    """

    # 更改欄位名稱為中文
    data = data.rename(columns={'date': '日期', 'stock_id': '股票代碼', 'Trading_Volume': '成交股數', 'Trading_money': '成交金額', 'open': '開盤價', 'max': '最高價', 'min': '最低價', 'close': '收盤價', 'spread': '漲跌', 'Trading_turnover': '成交筆數'})

    # 儲存 CSV 檔案
    csv_file = file_name + '.csv'
    if os.path.exists(csv_file):
        logger.warning(f"csv警告：{csv_file} 已存在，將被覆蓋")
    data.to_csv(csv_file, index=False)  # 添加 index=False

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    logger.debug(f"{ticker}_{identify_name}_{start_short}_{end_short} 資料數據如下")  # 細節用 DEBUG
    logger.info(f"\n{data.to_string(index=False)}")  # 無索引顯示

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data  # 返回處理後的 DataFrame，供後續使用

def taiwan_stock_margin_purchase_short_sale(ticker='2330', start_date='', end_date=''):
    """
    下載並處理台灣股市融資券數數據，繪製圖表並儲存 CSV。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'
        start_date (str): 開始日期，例如 '2025-02-17'
        end_date (str): 結束日期，例如 '2025-02-26'
        參考網址：https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?date=20250304&selectType=0099P&response=html
        API使用方式：https://finmind.github.io/tutor/TaiwanMarket/Chip/#taiwanstockmarginpurchaseshortsale
    """

    # 今天日期
    today = datetime.now().strftime('%Y-%m-%d')
    
    # 如果開始時間未定義，預設為今天往前減 30 天
    if not start_date:
        start_date = today
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        start_date = (start_date + timedelta(days=-30)).strftime('%Y-%m-%d')
    
    # 如果結束時間未定義，預設為今天，否則為原來的 end_date
    if not end_date:
        end_date_dt = datetime.strptime(today, '%Y-%m-%d')
    else:
        end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')

    end_date = (end_date_dt + timedelta(days=0)).strftime('%Y-%m-%d')

    # 檔名識別
    identify_name = 'finmind_margin_purchase'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯日期行為
    setup_check_realtime_format(logger, start_date, end_date, '>')

    # 簡化日期格式
    start_short = start_date.replace('-', '')
    end_short = end_date.replace('-', '')

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載股票融資券數數據
    try:
        data = dl.taiwan_stock_margin_purchase_short_sale(stock_id=ticker, start_date=start_date, end_date=end_date)
        data['date'] = pd.to_datetime(data['date'])

        # 只保留指定欄位
        keep_columns = [
            'date',
            'MarginPurchaseTodayBalance',
            'MarginPurchaseYesterdayBalance',
            'ShortSaleTodayBalance',
            'ShortSaleYesterdayBalance'
        ]
        data = data[keep_columns]  # 篩選指定欄位，其他欄位會被移除

        data = data.reset_index(drop=True)  # 移除數字索引
    except Exception as e:
        logger.error(f"下載 {ticker} 融資券數異常：{e}")
        return None  # 若發生連線異常，返回 None 並結束函數

    # 檢查並顯示數據
    if data.empty:
        logger.warning(f"沒有找到 {ticker} 的融資券數數據，可能是日期範圍無效或未來日期")
        return None  # 若無數據，返回 None 並結束函數

    # 格式化 融資融券 欄位，加入千位分隔符
    data['MarginPurchaseTodayBalance'] = data['MarginPurchaseTodayBalance'].apply(lambda x: f"{x:,}")
    data['MarginPurchaseYesterdayBalance'] = data['MarginPurchaseYesterdayBalance'].apply(lambda x: f"{x:,}")
    data['ShortSaleTodayBalance'] = data['ShortSaleTodayBalance'].apply(lambda x: f"{x:,}")
    data['ShortSaleYesterdayBalance'] = data['ShortSaleYesterdayBalance'].apply(lambda x: f"{x:,}")

    # 更改欄位名稱為中文
    data = data.rename(columns={'date': '日期', 'MarginPurchaseTodayBalance': '當日融資餘額', 'MarginPurchaseYesterdayBalance': '前一日融資餘額', 'ShortSaleTodayBalance': '當日融券餘額', 'ShortSaleYesterdayBalance': '前一日融券餘額'})

    # 儲存 CSV 檔案
    csv_file = file_name + '.csv'
    if os.path.exists(csv_file):
        logger.warning(f"csv警告：{csv_file} 已存在，將被覆蓋")
    data.to_csv(csv_file, index=False)  # 添加 index=False

    # 顯示資料
    logger.debug(f"{ticker}_margin_purchase_{start_short}_{end_short} 資料數據如下")  # 細節用 DEBUG
    logger.info(f"\n{data.to_string(index=False)}")  # 無索引顯示

    return data  # 返回處理後的 DataFrame，供後續使用

def taiwan_stock_institutional_investors(ticker='2330', start_date='', end_date=''):
    """
    下載並處理台灣股市法人買賣超數據，繪製圖表並儲存 CSV。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'
        start_date (str): 開始日期，例如 '2025-02-17'
        end_date (str): 結束日期，例如 '2025-02-26'
        參考網址：https://www.twse.com.tw/rwd/zh/fund/T86?date=20250306&selectType=0099P&response=html
        API使用方式：https://finmind.github.io/tutor/TaiwanMarket/Chip/#taiwanstockinstitutionalinvestorsbuysell
    """

    # 今天日期
    today = datetime.now().strftime('%Y-%m-%d')
    
    # 如果開始時間未定義，預設為今天往前減 30 天
    if not start_date:
        start_date = today
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        start_date = (start_date + timedelta(days=-30)).strftime('%Y-%m-%d')
    
    # 如果結束時間未定義，預設為今天，否則為原來的 end_date
    if not end_date:
        end_date_dt = datetime.strptime(today, '%Y-%m-%d')
    else:
        end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')

    end_date = (end_date_dt + timedelta(days=0)).strftime('%Y-%m-%d')

    # 檔名識別
    identify_name = 'finmind_institutional'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯日期行為
    setup_check_realtime_format(logger, start_date, end_date, '>')

    # 簡化日期格式
    start_short = start_date.replace('-', '')
    end_short = end_date.replace('-', '')

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載法人買賣超數據
    try:
        data = dl.taiwan_stock_institutional_investors(stock_id=ticker, start_date=start_date, end_date=end_date)
        data['date'] = pd.to_datetime(data['date'])
        
        # 移除指定欄位
        columns_to_drop = ['stock_id']  # 可擴展為多個欄位
        data = data.drop(columns=[col for col in columns_to_drop if col in data.columns])  # 只移除存在的欄位
        
        data = data.reset_index(drop=True)  # 移除數字索引
    except Exception as e:
        logger.error(f"下載 {ticker} 法人買賣超數據異常：{e}")
        return None  # 若發生連線異常，返回 None 並結束函數

    # 檢查並顯示數據
    if data.empty:
        logger.warning(f"沒有找到 {ticker} 的法人買賣超數據，可能是日期範圍無效或未來日期")
        return None  # 若無數據，返回 None 並結束函數

    # 格式化 buy 和 sell 和 net 欄位，加入千位分隔符
    data['buy'] = data['buy'].apply(lambda x: f"{x:,}")
    data['sell'] = data['sell'].apply(lambda x: f"{x:,}")
    data['net'] = pd.to_numeric(data['buy'].str.replace(',', '')) - pd.to_numeric(data['sell'].str.replace(',', ''))  # 轉為數字才能相減
    data['net'] = data['net'].apply(lambda x: f"{x:,}")  # 再補上千分位

    # 取代法人類型名稱為中文
    name_mapping = {
        'Foreign_Investor': '外陸資',
        'Foreign_Dealer_Self': '外資自營商',
        'Investment_Trust': '投信',
        'Dealer_self': '自營商（自行買賣）',
        'Dealer_Hedging': '自營商（避險）'
    }
    data['name'] = data['name'].replace(name_mapping)

    # 繪製分法人類型的折線圖
    plt.figure(figsize=(12, 6))  # 增加圖表寬度，避免擁擠
    colors = ['red', 'blue', 'green', 'orange', 'purple']  # 為五種法人指定顏色
    for i, (name, group) in enumerate(data.groupby('name')):
        net_values = pd.to_numeric(group['net'].str.replace(',', ''))
        plt.plot(group['date'], net_values, marker='o', linestyle='-', color=colors[i % len(colors)], label=name)
        for date, value in zip(group['date'], net_values):
            plt.text(date, value, f'{value:,.0f}', ha='left', va='bottom' if value >= 0 else 'top', fontsize=8)

    plt.title(f'{ticker} 法人淨買賣趨勢（按類型）')
    plt.grid(True)
    plt.xlabel('日期')
    plt.ylabel('淨買賣股數')
    plt.xticks(data['date'], rotation=90)
    plt.legend(title='法人類型')  # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角
    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 更改欄位名稱為中文
    data = data.rename(columns={'date': '日期', 'buy': '買進股數', 'sell': '賣出股數', 'name': '法人類型', 'net': '淨買賣'})

    # 儲存 CSV 檔案
    csv_file = file_name + '.csv'
    if os.path.exists(csv_file):
        logger.warning(f"csv警告：{csv_file} 已存在，將被覆蓋")
    data.to_csv(csv_file, index=False)  # 添加 index=False

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    logger.debug(f"{ticker}_{identify_name}_{start_short}_{end_short} 資料數據如下")  # 細節用 DEBUG
    logger.info(f"\n{data.to_string(index=False)}")  # 無索引顯示

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data  # 返回處理後的 DataFrame，供後續使用

def rsv(period=9, ticker='2330'):
    """
    計算指定股票的 RSV 指標（使用 yfinance）（Raw Stochastic Value），「加權移動平均」）。
    
    參數：
        period (int): 計算周期，例如 9
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
    """	

    # 檔名識別
    identify_name = 'yfinance_rsv'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查週期是否為數字
    try:
        period = int(period)
    except ValueError:
        logger.error(f"周期 {period} 必須為整數")
        return None

    # 下載數據（多抓一天以計算差值）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{period + 1}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{period + 1}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Open', 'High', 'Low', 'Close']]
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < period:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 {period} 天的 RSV")
        print(data)
        return None

    # 計算 RSV，確保取值為純量
    latest_close = data['Close'].iloc[-1].item()  # 最新收盤價
    period_low = data['Low'].min().item()         # 近 N 天最低價
    period_high = data['High'].max().item()       # 近 N 天最高價
    
    if period_high == period_low:
        rsv_value = 50.0  # 若最高等於最低，設為 50（避免除以 0）
    else:
        rsv_value = (latest_close - period_low) / (period_high - period_low) * 100

    # 輸出結果
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"近 {period} 天的最低價為：{period_low:.2f}\n"
        f"近 {period} 天的最高價為：{period_high:.2f}\n"
        f"RSV 為：{rsv_value:.2f}\n"
    )
    logger.info(result)

    return data

def kd(period=9, ticker='2330'):
    """
    計算指定股票的 KD 指標（使用 yfinance）（stochastic oscillator，又稱為隨機指標）。
    
    參數：
        period (int): 計算周期，例如 9 或 14
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
    """

    # 檔名識別
    identify_name = 'yfinance_kd'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查週期是否為數字
    try:
        period = int(period)
    except ValueError:
        logger.error(f"周期 {period} 必須為整數")
        return None

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}_{period}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（多抓一天以計算差值）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{period + 1}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{period + 1}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Open', 'High', 'Low', 'Close']]
        data['Date'] = pd.to_datetime(data['Date'])
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < period:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 {period} 天的 KD")
        return None

    # 計算 RSV
    data['RSV'] = 0.0
    for i in range(len(data)):
        if i < period - 1:
            data.loc[data.index[i], 'RSV'] = 50.0  # 前幾天無足夠數據，設為 50
        else:
            window = data.iloc[i - period + 1:i + 1]
            latest_close = data['Close'].iloc[i].item()
            period_low = window['Low'].min().item()
            period_high = window['High'].max().item()
            if period_high == period_low:
                data.loc[data.index[i], 'RSV'] = 50.0  # 若最高等於最低，設為 50（避免除以 0）
            else:
                data.loc[data.index[i], 'RSV'] = (latest_close - period_low) / (period_high - period_low) * 100
    
    # 計算 K 和 D 值
    data['K'] = 50.0  # 初始 K 值
    data['D'] = 50.0  # 初始 D 值
    for i in range(1, len(data)):
        prev_k = data['K'].iloc[i - 1]
        prev_d = data['D'].iloc[i - 1]
        current_rsv = data['RSV'].iloc[i]
        current_k = (2/3 * prev_k) + (1/3 * current_rsv)
        current_d = (2/3 * prev_d) + (1/3 * current_k)
        data.loc[data.index[i], 'K'] = current_k
        data.loc[data.index[i], 'D'] = current_d

    # 輸出最新 K 和 D
    latest_k = data['K'].iloc[-1]
    latest_d = data['D'].iloc[-1]
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"近 {period} 天的最低價為：{period_low:.2f}\n"
        f"近 {period} 天的最高價為：{period_high:.2f}\n"
        f"RSV 為：{current_rsv:.2f}\n"
        f"近 {period} 天的最新 K 值：{latest_k:.2f}\n"
        f"近 {period} 天的最新 D 值：{latest_d:.2f}\n"
        f"K > 80 可能超買，K < 20 可能超賣\n"
        f"K 如果一開始小於 D，顯示短期動能偏弱（下跌趨勢），但後來假設 K 慢慢接近 D，那就是所謂的景氣復甦，黃金交叉，且價格同步上漲，可能是買入訊號\n"
        f"K 如果一開始大於 D，顯示短期動能偏強（上漲趨勢），但後來假設 K 慢慢接近 D，那就是所謂的景氣衰退，死亡交叉，且價格同步下跌，可能是賣出訊號\n"
        f"若 K 持續低於 D，且差距擴大，可能賣壓加重\n"
        f"若 K 持續高於 D，且差距擴大，可能漲勢強勁\n"
    )
    logger.info(result)

    # 儲存當日 KD 到歷史記錄（四捨五入到 2 位小數）
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')  # 轉為字符串
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'Period_Low': [round(period_low, 2)],
        'Period_High': [round(period_high, 2)],
        'RSV': [round(current_rsv, 2)],
        'K': [round(latest_k, 2)],
        'D': [round(latest_d, 2)]
    })

    history_file = f'output/{ticker}/{ticker}_{identify_name}_{period}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)  # 首次寫入，帶表頭
    else:
        existing_data = pd.read_csv(history_file)  # 不設 index_col
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)  # 追加，不帶表頭
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])  # 轉為 datetime

    # 繪製 KD 線
    plt.figure(figsize=(12, 6))
    plt.plot(history_data_load['Date'], history_data_load['K'], marker='o', linestyle='-', color='blue', label='K 值')
    plt.plot(history_data_load['Date'], history_data_load['D'], marker='o', linestyle='-', color='red', label='D 值')
    plt.axhline(y=80, color='gray', linestyle='--', label='超買線（k > 80）')
    plt.axhline(y=20, color='gray', linestyle='--', label='超賣線（k < 20）')

    # 標示每個點的數值
    for date, k, d in zip(history_data_load['Date'], history_data_load['K'], history_data_load['D']):
        plt.text(date, k, f'{k:.2f}', ha='left', va='bottom' if k >= d else 'top', fontsize=8, color='blue')
        plt.text(date, d, f'{d:.2f}', ha='right', va='bottom' if d >= k else 'top', fontsize=8, color='red')

    plt.title(f'{ticker} KD 指標（歷史趨勢）')
    plt.xlabel('日期')
    plt.ylabel('數值')
    plt.grid(True)
    plt.xticks(history_data_load['Date'], rotation=90)
    plt.legend() # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角
    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    history_data_load = pd.read_csv(history_file)  # 再讀取一次最新的資料，因為可能會有新寫入的資料面
    logger.debug(f"{ticker}_{identify_name}_{period} 完整數據如下")  # 細節用 DEBUG
    logger.info(f"\n{history_data_load.to_string(index=False)}")

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data  # 返回處理後的 DataFrame，供後續使用

def ma(ticker='2330', short_period=5, long_period=200):
    """
    分析股票的移動平均線，又稱為均線（Moving Average）。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
        short_period (int): 短期 MA 周期，例如 5 天（週線）
        long_period (int): 長期 MA 周期，例如 200 天（年線）
    """
    
    # 檔名識別
    identify_name = 'yfinance_ma'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查週期是否為數字
    try:
        short_period = int(short_period)
        long_period = int(long_period)
    except ValueError:
        logger.error(f"周期 {short_period} 或 {long_period} 必須為整數")
        return None

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}_{short_period}_{long_period}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（多抓一些天數以計算長期 MA）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{long_period + 10}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{long_period + 10}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Close']]
        data['Date'] = pd.to_datetime(data['Date'])
        logger.debug(f"下載數據長度：{len(data)} 天，從 {data['Date'].iloc[0]} 到 {data['Date'].iloc[-1]}")
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    # 計算不同周期的 MA
    data['MA5'] = data['Close'].rolling(window=short_period, min_periods=1).mean()
    data['MA10'] = data['Close'].rolling(window=10, min_periods=1).mean()
    data['MA20'] = data['Close'].rolling(window=20, min_periods=1).mean()
    data['MA60'] = data['Close'].rolling(window=60, min_periods=1).mean()
    data['MA120'] = data['Close'].rolling(window=120, min_periods=1).mean()
    data['MA200'] = data['Close'].rolling(window=long_period, min_periods=1).mean()

    # 取最近一段數據（30 天）
    plot_data = data.tail(30)
    logger.debug(f"繪圖數據長度：{len(plot_data)} 天，從 {plot_data['Date'].iloc[0]} 到 {plot_data['Date'].iloc[-1]}")

    # 輸出最新股價與 MA
    latest_close = data['Close'].iloc[-1].item()
    latest_ma5 = data['MA5'].iloc[-1].item()
    latest_ma10 = data['MA10'].iloc[-1].item()
    latest_ma20 = data['MA20'].iloc[-1].item()
    latest_ma60 = data['MA60'].iloc[-1].item()
    latest_ma120 = data['MA120'].iloc[-1].item()
    latest_ma200 = data['MA200'].iloc[-1].item()
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"{short_period} 天移動平均線：{latest_ma5:.2f}\n"
        f"10 天移動平均線：{latest_ma10:.2f}\n"
        f"20 天移動平均線：{latest_ma20:.2f}\n"
        f"60 天移動平均線：{latest_ma60:.2f}\n"
        f"120 天移動平均線：{latest_ma120:.2f}\n"
        f"{long_period} 天移動平均線：{latest_ma200:.2f}\n"
        f"股價高於 MA 可能為買入訊號，低於可能為賣出訊號\n"
        f"股價上穿 MA 可能為買入訊號，下穿可能為賣出訊號\n"
    )
    logger.info(result)
    
    # 儲存當日 MA 到歷史記錄
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'latest_ma5': [round(latest_ma5, 2)],
        'latest_ma10': [round(latest_ma10, 2)],
        'latest_ma20': [round(latest_ma20, 2)],
        'latest_ma60': [round(latest_ma60, 2)],
        'latest_ma120': [round(latest_ma120, 2)],
        'latest_ma200': [round(latest_ma200, 2)]
    })
    history_file = f'output/{ticker}/{ticker}_{identify_name}_{short_period}_{long_period}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)
    else:
        existing_data = pd.read_csv(history_file)
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 繪製圖表
    plt.figure(figsize=(12, 6))
    plt.plot(plot_data['Date'], plot_data['Close'], marker='o', linestyle='-', color='black', label=f'歷史股價（目前為：{latest_close:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA5'], linestyle='-', color='green', label=f'{short_period} 天 MA（週線）（{latest_ma5:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA10'], linestyle='-', color='blue', label=f'10 天 MA（雙週線）（{latest_ma10:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA20'], linestyle='-', color='purple', label=f'20 天 MA（月線）（{latest_ma20:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA60'], linestyle='-', color='orange', label=f'60 天 MA（季線）（{latest_ma60:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA120'], linestyle='-', color='red', label=f'120 天 MA（半年線）（{latest_ma120:.2f}）')
    plt.plot(plot_data['Date'], plot_data['MA200'], linestyle='-', color='brown', label=f'{long_period} 天 MA（年線）（{latest_ma200:.2f}）')

    # 標示股價與 MA 交叉點（台灣股市：紅色上漲，綠色下跌）
    for ma_col, color in [('MA5', 'green'), ('MA10', 'blue'), ('MA20', 'purple'), ('MA60', 'orange'), ('MA120', 'red'), ('MA200', 'brown')]:
        for i in range(1, len(plot_data)):
            prev_close = plot_data['Close'].iloc[i-1].item()
            curr_close = plot_data['Close'].iloc[i].item()
            prev_ma = plot_data[ma_col].iloc[i-1].item()
            curr_ma = plot_data[ma_col].iloc[i].item()
            if prev_close < prev_ma and curr_close > curr_ma:  # 上穿
                plt.scatter(plot_data['Date'].iloc[i], curr_close, color='red', marker='^', s=100)
            elif prev_close > prev_ma and curr_close < curr_ma:  # 下穿
                plt.scatter(plot_data['Date'].iloc[i], curr_close, color='green', marker='v', s=100)

    # 標示股價數值
    logger.debug(f"plot_data 結構：{plot_data.to_string()}")
    for i in range(len(plot_data)):
        date = plot_data['Date'].iloc[i]
        close_value = plot_data['Close'].iloc[i].item()
        plt.text(date, close_value, f'{close_value:.2f}', ha='left', va='bottom' if close_value > plot_data['MA120'].mean() else 'top', fontsize=8, color='black')

    plt.title(f'{ticker} 股價與移動平均線')
    plt.xlabel('日期')
    plt.ylabel('價格')
    plt.grid(True)
    plt.xticks(plot_data['Date'], rotation=90)
    plt.xlim(plot_data['Date'].iloc[0], plot_data['Date'].iloc[-1])  # 限制 X 軸範圍
    plt.legend()  # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角
    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 儲存圖表
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data

def rsi(period=14, ticker='2330'):
    """
    計算指定股票的 RSI 指標（使用 yfinance）（Relative Strength Index，「相對強弱指標」）。
    
    參數：
        period (int): 計算周期，例如 14（市場標準）
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
    """

    # 檔名識別
    identify_name = 'yfinance_rsi'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查週期是否為數字
    try:
        period = int(period)
    except ValueError:
        logger.error(f"周期 {period} 必須為整數")
        return None

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}_{period}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（多抓一天以計算差值）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{period + 1}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{period + 1}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Close']]
        data['Date'] = pd.to_datetime(data['Date'])
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < period + 1:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 {period} 天的 RSI")
        return None
        
    # 計算每日價格變動
    data['Change'] = data['Close'].diff()
    data['Gain'] = data['Change'].apply(lambda x: x if x > 0 else 0)
    data['Loss'] = data['Change'].apply(lambda x: -x if x < 0 else 0)

    # 計算平均上升和下跌（使用前 period 天初始化）
    data['Avg_Gain'] = data['Gain'].rolling(window=period, min_periods=1).mean()
    data['Avg_Loss'] = data['Loss'].rolling(window=period, min_periods=1).mean()

    # 平滑 RSI（從 period+1 天開始）
    for i in range(period + 1, len(data)):
        data.loc[data.index[i], 'Avg_Gain'] = ((data['Avg_Gain'].iloc[i-1] * (period - 1)) + data['Gain'].iloc[i]) / period
        data.loc[data.index[i], 'Avg_Loss'] = ((data['Avg_Loss'].iloc[i-1] * (period - 1)) + data['Loss'].iloc[i]) / period

    # 計算 RSI
    data['RS'] = data['Avg_Gain'] / data['Avg_Loss'].replace(0, 0.0001)  # 避免除以 0
    data['RSI'] = 100 - (100 / (1 + data['RS']))

    # 輸出最新 RSI
    latest_close = data['Close'].iloc[-1].item()
    latest_rsi = data['RSI'].iloc[-1].item()
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"近 {period} 天的最新 RSI 值：{latest_rsi:.2f}\n"
        f"RSI > 70 可能超買，RSI < 30 可能超賣\n"
    )
    logger.info(result)

    # 儲存當日 RSI 到歷史記錄
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'RSI': [round(latest_rsi, 2)]
    })
    history_file = f'output/{ticker}/{ticker}_{identify_name}_{period}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)
    else:
        existing_data = pd.read_csv(history_file)
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])

    # 繪製 RSI 線
    plt.figure(figsize=(12, 6))
    plt.plot(history_data_load['Date'], history_data_load['RSI'], marker='o', linestyle='-', color='purple', label='RSI 值')
    plt.axhline(y=70, color='gray', linestyle='--', label='超買線（RSI > 70）')
    plt.axhline(y=30, color='gray', linestyle='--', label='超賣線（RSI < 30）')

    # 標示每個點的數值
    for date, rsi in zip(history_data_load['Date'], history_data_load['RSI']):
        plt.text(date, rsi, f'{rsi:.2f}', ha='left', va='bottom' if rsi >= 50 else 'top', fontsize=8, color='purple')

    plt.title(f'{ticker} RSI 指標（歷史趨勢，週期：{period}）')
    plt.xlabel('日期')
    plt.ylabel('數值')
    plt.grid(True)
    plt.xticks(history_data_load['Date'], rotation=90)
    plt.legend()  # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角
    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 儲存圖表為 PNG
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    history_data_load = pd.read_csv(history_file)
    logger.debug(f"{ticker}_{identify_name}_{period} 完整數據如下")
    logger.info(f"\n{history_data_load.to_string(index=False)}")

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data

def macd(ticker='2330', fast_period=12, slow_period=26, signal_period=9):
    """
    計算指定股票的 MACD 指標（使用 yfinance）（Moving Average Convergence & Divergence，平滑異同移動平均線指標）「平滑異同移動平均線指標」）。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
        fast_period (int): 短期 EMA 周期，例如 12
        slow_period (int): 長期 EMA 周期，例如 26
        signal_period (int): 訊號線 EMA 周期，例如 9
    """

    # 檔名識別
    identify_name = 'yfinance_macd'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查週期是否為數字
    try:
        fast_period = int(fast_period)
        slow_period = int(slow_period)
        signal_period = int(signal_period)
    except ValueError:
        logger.error(f"周期 {fast_period}, {slow_period}, 或 {signal_period} 必須為整數")
        return None

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}_{fast_period}_{slow_period}_{signal_period}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（多抓一些天數以穩定 EMA）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{slow_period + signal_period}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{slow_period + signal_period}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Close']]
        data['Date'] = pd.to_datetime(data['Date'])
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < slow_period + signal_period:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 MACD({fast_period},{slow_period},{signal_period})")
        return None

    # 計算 EMA 和 MACD
    data['EMA_Fast'] = data['Close'].ewm(span=fast_period, adjust=False).mean()
    data['EMA_Slow'] = data['Close'].ewm(span=slow_period, adjust=False).mean()
    data['MACD'] = data['EMA_Fast'] - data['EMA_Slow']
    data['Signal'] = data['MACD'].ewm(span=signal_period, adjust=False).mean()
    data['Histogram'] = data['MACD'] - data['Signal']

    # 輸出最新 MACD
    latest_close = data['Close'].iloc[-1].item()
    latest_macd = data['MACD'].iloc[-1].item()
    latest_signal = data['Signal'].iloc[-1].item()
    latest_histogram = data['Histogram'].iloc[-1].item()
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"近 {fast_period},{slow_period},{signal_period} 天的最新 MACD 值：{latest_macd:.2f}\n"
        f"訊號線值：{latest_signal:.2f}\n"
        f"柱狀圖值：{latest_histogram:.2f}\n"
        f"MACD 上穿訊號線（黃金交叉）可能為買入訊號，下穿（死亡交叉）可能為賣出訊號\n"
        f"柱狀圖 > 0 且增加表示上漲動能增強，< 0 且減少表示下跌動能增強"
    )
    logger.info(result)

    # 儲存當日 MACD 到歷史記錄
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'MACD': [round(latest_macd, 2)],
        'Signal': [round(latest_signal, 2)],
        'Histogram': [round(latest_histogram, 2)]
    })
    history_file = f'output/{ticker}/{ticker}_{identify_name}_{fast_period}_{slow_period}_{signal_period}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)
    else:
        existing_data = pd.read_csv(history_file)
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])

    # 繪製 MACD 圖表
    plt.figure(figsize=(12, 8))
    plt.subplot(2, 1, 1)
    plt.plot(history_data_load['Date'], history_data_load['MACD'], marker='o', linestyle='-', color='blue', label='MACD 線')
    plt.plot(history_data_load['Date'], history_data_load['Signal'], marker='o', linestyle='-', color='red', label='訊號線')
    plt.axhline(y=0, color='gray', linestyle='--')
    
    # 標註 MACD 和訊號線數值
    for date, macd, signal in zip(history_data_load['Date'], history_data_load['MACD'], history_data_load['Signal']):
        plt.text(date, macd, f'{macd:.2f}', ha='left', va='bottom' if macd >= signal else 'top', fontsize=8, color='blue')
        plt.text(date, signal, f'{signal:.2f}', ha='right', va='bottom' if signal >= macd else 'top', fontsize=8, color='red')

    plt.title(f'{ticker} MACD 指標（歷史趨勢，{fast_period},{slow_period},{signal_period}）')
    plt.xlabel('日期')
    plt.ylabel('數值')
    plt.grid(True)
    plt.xticks(history_data_load['Date'], rotation=90)
    plt.legend()  # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角

    plt.subplot(2, 1, 2)
    plt.bar(history_data_load['Date'], history_data_load['Histogram'], color='gray', label='柱狀圖')
    plt.axhline(y=0, color='black', linestyle='--')
    # 標註柱狀圖數值
    for date, hist in zip(history_data_load['Date'], history_data_load['Histogram']):
        plt.text(date, hist, f'{hist:.2f}', ha='center', va='bottom' if hist >= 0 else 'top', fontsize=8, color='black')
    plt.xlabel('日期')
    plt.ylabel('柱狀圖')
    plt.grid(True)
    plt.xticks(history_data_load['Date'], rotation=90)
    plt.legend()  # 添加圖例，若有需要可使用 loc='upper right' ，讓其常駐在右上角

    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 儲存圖表為 PNG
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    history_data_load = pd.read_csv(history_file)
    logger.debug(f"{ticker}_{identify_name}_{fast_period}_{slow_period}_{signal_period} 完整數據如下")
    logger.info(f"\n{history_data_load.to_string(index=False)}")

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data

def bollinger(ticker='2330', period=20, multiplier=2):
    """
    計算指定股票的布林通道（Bollinger Bands），並繪製圖表。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'（會自動加上 .TW）
        period (int): 計算周期，例如 20（市場標準）
        multiplier (float): 標準差倍數，例如 2（市場標準）
    """

    # 檔名識別
    identify_name = 'yfinance_bollinger'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 檢查參數是否有效
    try:
        period = int(period)
        multiplier = float(multiplier)
    except ValueError:
        logger.error(f"周期 {period} 必須為整數，倍數 {multiplier} 必須為數字")
        return None

    if period <= 0 or multiplier <= 0:
        logger.error(f"周期 {period} 和倍數 {multiplier} 必須大於 0")
        return None

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}_{period}_{multiplier}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（多抓一些天數以穩定計算）
    try:
        # data = yf.download(f"{ticker}.TW", period=f'{period + 10}d')
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period=f'{period + 10}d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        # data = data[['Date', 'Close']]
        data['Date'] = pd.to_datetime(data['Date'])
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < period:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 {period} 天的布林通道")
        return None

    # 計算布林通道
    data['Middle_Band'] = data['Close'].rolling(window=period, min_periods=1).mean()  # 中軌：SMA
    data['Std_Dev'] = data['Close'].rolling(window=period, min_periods=1).std()      # 標準差
    data['Upper_Band'] = data['Middle_Band'] + (multiplier * data['Std_Dev'])        # 上軌
    data['Lower_Band'] = data['Middle_Band'] - (multiplier * data['Std_Dev'])        # 下軌

    # 輸出最新布林通道數據
    latest_close = data['Close'].iloc[-1].item()
    latest_middle = data['Middle_Band'].iloc[-1].item()
    latest_upper = data['Upper_Band'].iloc[-1].item()
    latest_lower = data['Lower_Band'].iloc[-1].item()
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"近 {period} 天的布林通道中軌（SMA）：{latest_middle:.2f}\n"
        f"上軌（{multiplier} 倍標準差）：{latest_upper:.2f}\n"
        f"下軌（{multiplier} 倍標準差）：{latest_lower:.2f}\n"
        f"股價接近或突破上軌可能表示超買\n"
        f"股價接近或跌破下軌可能表示超賣\n"
        f"通道收窄可能預示即將出現大波動"
    )
    logger.info(result)

    # 儲存當日布林通道到歷史記錄
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'Middle_Band': [round(latest_middle, 2)],
        'Upper_Band': [round(latest_upper, 2)],
        'Lower_Band': [round(latest_lower, 2)]
    })
    history_file = f'output/{ticker}/{ticker}_{identify_name}_{period}_{multiplier}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)
    else:
        existing_data = pd.read_csv(history_file)
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])

    # 繪製布林通道圖表
    plt.figure(figsize=(12, 6))
    plt.plot(history_data_load['Date'], history_data_load['Close'], marker='o', linestyle='-', color='black', label='股價')
    plt.plot(history_data_load['Date'], history_data_load['Middle_Band'], linestyle='-', color='blue', label=f'中軌 ({period} 天 SMA)')
    plt.plot(history_data_load['Date'], history_data_load['Upper_Band'], linestyle='--', color='red', label=f'上軌 ({multiplier} 倍標準差)')
    plt.plot(history_data_load['Date'], history_data_load['Lower_Band'], linestyle='--', color='green', label=f'下軌 ({multiplier} 倍標準差)')
    plt.fill_between(history_data_load['Date'], history_data_load['Lower_Band'], history_data_load['Upper_Band'], 
                     color='gray', alpha=0.2, label='布林通道範圍')

    # 標示所有線段的數值
    for date, close, middle, upper, lower in zip(history_data_load['Date'], 
                                                 history_data_load['Close'], 
                                                 history_data_load['Middle_Band'], 
                                                 history_data_load['Upper_Band'], 
                                                 history_data_load['Lower_Band']):
        # 標示股價（黑色）
        plt.text(date, close, f'{close:.2f}', ha='left', va='bottom' if close >= middle else 'top', 
                 fontsize=8, color='black')
        # 標示中軌（藍色）
        plt.text(date, middle, f'{middle:.2f}', ha='left', va='bottom' if middle > close else 'top', 
                 fontsize=8, color='blue')
        # 標示上軌（紅色）
        plt.text(date, upper, f'{upper:.2f}', ha='left', va='bottom', 
                 fontsize=8, color='red')
        # 標示下軌（綠色）
        plt.text(date, lower, f'{lower:.2f}', ha='left', va='top', 
                 fontsize=8, color='green')

    plt.title(f'{ticker} 布林通道（週期：{period}，倍數：{multiplier}）')
    plt.xlabel('日期')
    plt.ylabel('價格')
    plt.grid(True)
    plt.xticks(history_data_load['Date'], rotation=90)
    plt.legend()  # 添加圖例，若有需要可使用 loc='upper right'
    plt.tight_layout()  # 自動調整佈局，避免標籤被裁切

    # 儲存圖表為 PNG
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示資料
    history_data_load = pd.read_csv(history_file)
    logger.debug(f"{ticker}_{identify_name}_{period}_{multiplier} 完整數據如下")
    logger.info(f"\n{history_data_load.to_string(index=False)}")

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data

def obv(ticker='2330'):
    """
    計算指定股票的能量潮指標（On Balance Volume, OBV），並繪製圖表。
    如果當日收盤價 > 前一日收盤價，當日成交量（Volume）以正值加入 OBV。
    如果當日收盤價 < 前一日收盤價，當日成交量以負值加入 OBV（即減去成交量）。
    如果當日收盤價 = 前一日收盤價，OBV 不變（成交量不計入）。

    成交量與價格的背離：
        如果價格上漲，但 OBV 下跌（或持平），可能表示上漲缺乏成交量支持，是假突破。
        如果價格下跌，但 OBV 上漲（或持平），可能表示下跌是洗盤，買盤暗中累積。
        這種背離能幫助判斷趨勢的真實性。
    趨勢的強度：
        OBV 持續上升，配合價格上漲，說明趨勢強勁且有成交量支撐。
        OBV 持續下跌，配合價格下跌，說明賣壓強大。
    長期累積效應：
    OBV 適合用來觀察大資金的進出。例如，若 OBV 在某段期間從負值轉為正值，可能暗示市場情緒從悲觀轉為樂觀。
    
    參數：
        ticker (str): 股票代碼，例如 '2330'（會自動加上 .TW）
    """

    # 檔名識別
    identify_name = 'yfinance_obv'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯資料夾行為
    output_dir = setup_folders(ticker)

    # 設定檔名統一命名方式
    file_name = f'{output_dir}/{ticker}_{identify_name}'

    # 關聯字型行為
    setup_chinese_font(logger)

    # 下載數據（預設抓取最近 30 天數據）
    try:
        stock = yf.Ticker(f"{ticker}.TW")
        data = stock.history(period='30d', auto_adjust=False)
        if data.empty:
            logger.warning(f"沒有找到 {ticker}.TW 的數據，可能是無效代碼或無交易日")
            return None
        data = data.reset_index()
        data['Date'] = pd.to_datetime(data['Date'])
    except Exception as e:
        logger.error(f"下載 {ticker}.TW 數據異常：{e}")
        return None

    if len(data) < 2:
        logger.warning(f"數據不足，只有 {len(data)} 天，無法計算 OBV")
        return None

    # 計算 OBV
    data['Price_Change'] = data['Close'].diff()  # 計算收盤價變化
    data['OBV'] = 0  # 初始化 OBV 列
    for i in range(1, len(data)):
        if data['Price_Change'].iloc[i] > 0:
            data.loc[i, 'OBV'] = data['Volume'].iloc[i]  # 收盤價上升，成交量為正
        elif data['Price_Change'].iloc[i] < 0:
            data.loc[i, 'OBV'] = -data['Volume'].iloc[i]  # 收盤價下降，成交量為負
        else:
            data.loc[i, 'OBV'] = 0
    data['OBV'] = data['OBV'].cumsum()  # 累加計算 OBV
    data['OBV_Change'] = data['OBV'].diff()  # 新增 OBV 日變動

    # 輸出最新 OBV 數據與解釋
    latest_close = data['Close'].iloc[-1].item()
    latest_obv = data['OBV'].iloc[-1].item()
    latest_obv_change = data['OBV_Change'].iloc[-1].item()
    result = (
        f"目前收盤價為：{latest_close:.2f}\n"
        f"最新 OBV 值：{latest_obv:.0f}\n"
        f"最新 OBV_Change 值：{latest_obv_change:.0f}\n"
        f"OBV 上漲表示成交量支持價格上升趨勢\n"
        f"OBV 下跌表示成交量支持價格下降趨勢\n"
        f"OBV 趨平或震盪可能表示市場混亂，需謹慎觀察"
    )
    logger.info(result)

    # 判斷 Description
    last_price_change = data['Price_Change'].iloc[-1]
    last_obv_change = data['OBV_Change'].iloc[-1] if not pd.isna(data['OBV_Change'].iloc[-1]) else 0
    if last_price_change < 0 and last_obv_change > 0:
        description = "洗盤或暗中交易"
    elif last_price_change > 0 and last_obv_change < 0:
        description = "假突破"
    else:
        description = "無特別"

    # 儲存當日 OBV 到歷史記錄
    last_date_str = data['Date'].iloc[-1].strftime('%Y-%m-%d')
    history_data = pd.DataFrame({
        'Date': [last_date_str],
        'Close': [round(latest_close, 2)],
        'OBV': [round(latest_obv, 0)],
        'Description': [description]  # 新增 Description 欄位
    })
    history_file = f'output/{ticker}/{ticker}_{identify_name}.csv'
    if not os.path.exists(history_file):
        history_data.to_csv(history_file, index=False)
    else:
        existing_data = pd.read_csv(history_file)
        logger.debug(f"檢查日期：last_date={last_date_str}, existing_dates={existing_data['Date'].tolist()}")
        if last_date_str not in existing_data['Date'].astype(str).values:
            logger.info(f"追加新數據：{last_date_str}")
            history_data.to_csv(history_file, index=False, mode='a', header=False)
        else:
            logger.info(f"日期 {last_date_str} 已存在，跳過寫入")

    # 讀取歷史數據
    history_data_load = pd.read_csv(history_file)
    history_data_load['Date'] = pd.to_datetime(history_data_load['Date'])

    # 繪製折線圖，明確指定 X 軸數據
    plt.plot(history_data_load['Date'], history_data_load['Close'], marker='o', linestyle='-', color='blue', label='股價')
    plt.title(f'{ticker} 能量潮指標（OBV）')
    plt.grid(True)  # 添加網格線
    plt.xlabel('日期')
    plt.legend(title='價格類型')
    plt.xticks(history_data_load['Date'], rotation=90)

    # 在每個圓點旁顯示值 - close 和 Description
    for i, (date, close, obv, desc) in enumerate(zip(history_data_load['Date'], history_data_load['Close'], history_data_load['OBV'], history_data_load['Description'])):
        label = f'{close:.2f} ({desc})'  # 改成顯示 Description
        va = 'bottom' if obv >= 0 else 'top'
        plt.text(date, close, label, ha='left', va=va, fontsize=8, 
                 color='red' if obv >= 0 else 'green')

    # 儲存圖表為 PNG 檔案
    png_file = file_name + '.png'
    if os.path.exists(png_file):
        logger.warning(f"png警告：{png_file} 已存在，將被覆蓋")
    plt.savefig(png_file)

    # 顯示圖表並釋放記憶體資源
    # plt.show()
    plt.close()

    return data  # 返回處理後的 DataFrame，供後續使用

def etf_data(ticker='2330'):
    """
    爬蟲etf股票數據，繪製圖表並儲存 CSV。
    
    參數：
        ticker (str): 股票代碼，例如 '0052'
    """
    
    # 檔名識別
    identify_name = 'etf_data'
    
    # 關聯log行為
    logger = setup_logger(ticker, identify_name)
    
    func_name = 'fetch_' + ticker  # 動態生成函數名稱，例如 'fetch_0052'
    if func_name in globals():
        return globals()[func_name](logger, ticker, identify_name)  # 呼叫對應函數
    else:
        raise ValueError(f"沒有找到 {func_name} 的函數")

def fetch_0052(logger=None, ticker='2330', identify_name=''):
    url = "https://websys.fsit.com.tw/FubonETF/Trade/Estimate.aspx?no=Domestic&lan=TW&sty=BOX"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        # 發送請求並確保請求成功
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 解析 HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 找到目標 div
        div = soup.find('div', class_='con_c1_table_con')
        if not div:
            logger.error("無法找到 class='con_c1_table_con' 的 div")
            return None
        
        # 找到 0052 的 tr
        target_tr = None
        for tr in div.find_all('tr'):
            td_name = tr.find('td', class_='name_txt blue1 bluebg2')
            if td_name and '0052' in td_name.text:
                target_tr = tr
                break
        
        if not target_tr:
            logger.error("無法找到 0052 的資料行")
            return None
        
        # 提取所有 td
        tds = target_tr.find_all('td')
        if len(tds) < 10:
            logger.error("資料行欄位數不足")
            return None
        
        # 提取並格式化 timestamp
        raw_timestamp = tds[9].text.strip()  # 原始數據，例如 "2025/03/0717:05:00"
        formatted_timestamp = f"{raw_timestamp[:10]} {raw_timestamp[10:]}"  # 加入空白，變成 "2025/03/07 17:05:00"
        
        # 提取目標數據
        data = {
            'yesterday_nav': float(tds[1].find('span', class_='blue1').text),  # 昨天淨值
            'current_nav': float(tds[2].text),  # 當下淨值
            'yesterday_price': float(tds[4].find('span', class_='blue1').text),  # 昨天股價
            'current_price': float(tds[5].text),  # 當下股價
            'timestamp': formatted_timestamp  # 格式化後的時間
        }
        history_data = pd.DataFrame([data])
        history_file = f'output/{ticker}/{ticker}_{identify_name}_nav.csv'
        if not os.path.exists(history_file):
            history_data.to_csv(history_file, index=False)  # 首次寫入，帶表頭
        else:
            existing_data = pd.read_csv(history_file)  # 不設 index_col
            logger.debug(f"檢查日期：last_date={formatted_timestamp}, existing_dates={existing_data['timestamp'].tolist()}")
            if formatted_timestamp not in existing_data['timestamp'].astype(str).values:
                logger.info(f"追加新數據：{formatted_timestamp}")
                history_data.to_csv(history_file, index=False, mode='a', header=False)  # 追加，不帶表頭
            else:
                logger.info(f"日期 {formatted_timestamp} 已存在，跳過寫入")
        
        logger.info(f"成功抓取 0052 數據：{data}")
        return data
    
    except requests.RequestException as e:
        logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        logger.error(f"數據解析錯誤：{str(e)}")
        return None

def fetch_0050(logger=None):
    return "抓取 0050 的數據"

# GUI 相關的全局變數
window_alive = True
entries = []

def calculate_confidence_score(yesterday_close, avg_price, inner_vol, outer_vol, opening_vol, low_price, high_price, open_price, bid_vol, ask_vol, target_price, current_price, index_current_price, index_inner_vol, index_outer_vol, index_low_price, index_high_price, index_open_price):
    """
    初始分數：score = 0.5，基礎分數，設為中性起點（滿分 1.0，若未來新增新條件，初始分數保持 0.5，總加分仍為 0.5，需按比重重新分配各項分數。）。
    理由列表：reasons 用來記錄每個調整的邏輯依據，供 GUI 顯示。
    """
    reasons = []
    score = 0.5

    """
    邏輯：計算外盤量（買進成交量）佔總成交量（內盤 + 外盤）的比例。
        09:00-11:00：開盤量依價格方向計入外盤或內盤，影響上午信心。
        11:00-13:30：僅用盤中內外盤，反映消息與國際影響。
        pressure_ratio > 0.55：外盤佔比超過 55%，表示買盤力量強，分數加 0.15。
        pressure_ratio < 0.45：外盤佔比低於 45%，賣壓較重，分數減 0.15。
    意圖：內外盤反映市場短期買賣力量，外盤強意味買氣旺盛，提升信心。
    防呆：分母為 0 時，設為 0，避免除以零錯誤。
    """
    # 1. 內外盤壓力比
    # -1. 總加分上限：0.15

    # 自動偵測當前時間
    current_time = datetime.now().time()

    # 定義上午時段
    morning_start = time(9, 0)  # 09:00
    morning_end = time(11, 0)   # 11:00
    
    if morning_start <= current_time <= morning_end:
        # 上午：開盤量依價格方向調整
        if open_price > yesterday_close:
            outer_vol += opening_vol
            reasons.append(f"開盤價 {open_price} > 前日收盤 {yesterday_close}，開盤量 {opening_vol} 計入外盤")
        elif open_price < yesterday_close:
            inner_vol += opening_vol
            reasons.append(f"開盤價 {open_price} < 前日收盤 {yesterday_close}，開盤量 {opening_vol} 計入內盤")
        # 若相等，開盤量不調整，保持原始內外盤
    else:
        # 下午：不納入開盤量，使用原始內外盤
        reasons.append("下午時段，僅使用盤中內外盤數據")

    pressure_ratio = outer_vol / (inner_vol + outer_vol) if (inner_vol + outer_vol) > 0 else 0
    if pressure_ratio > 0.55:
        score += 0.15
        reasons.append(f"外盤量 {outer_vol} 佔比 {pressure_ratio:.2%} 大於內盤量 {inner_vol}，買盤偏強")
    elif pressure_ratio < 0.45:
        score -= 0.15
        reasons.append(f"內盤量 {inner_vol} 佔比 {1-pressure_ratio:.2%} 小於外盤量 {outer_vol}，賣壓稍重")

    """
    邏輯：計算委買量與委賣量的相對差異（第一檔）。
        imbalance = (bid_vol - ask_vol) / (bid_vol + ask_vol)：正值表示委買強，負值表示委賣強。
        imbalance > 0.2：委買量明顯高於委賣量（20% 以上），加 0.15。
        imbalance < -0.2：委賣量明顯高於委買量，減 0.05。
        範圍 -0.2 ~ 0.2：視為平衡，不調整。
    意圖：委買委賣的不平衡反映即時供需，委買強可能有支撐力。
    調整：權重較低（0.15/0.05），因為只看第一檔數據，影響不如內外盤全面。
    防呆：分母為 0 時設為 0。
    """
    # 2. 訂單簿不平衡（只看正負1檔）
    # -1. 總加分上限：0.1
    imbalance = (bid_vol - ask_vol) / (bid_vol + ask_vol) if (bid_vol + ask_vol) > 0 else 0
    if imbalance > 0.2:  # 提高門檻，因單檔數據更敏感
        score += 0.1  # 降低權重
        reasons.append(f"委買量 {bid_vol} 高於委賣量 {ask_vol}，不平衡度 {imbalance:.2f}，支撐偏強")
    elif imbalance < -0.2:
        score -= 0.1
        reasons.append(f"委賣量 {ask_vol} 高於委買量 {bid_vol}，不平衡度 {imbalance:.2f}，壓力略大")

    """
    邏輯：
        range_position：目標價格在當日價格區間（低點到高點）的相對位置（0 ~ 1）。
        open_position：目標價格相對開盤價的變化率。
        range_position < 0.4 or open_position < -0.005：目標價格接近低點（低於 40% 區間）或低於開盤價 0.5%，加 0.15。
        range_position > 0.6：目標價格接近高點（高於 60% 區間），減 0.05。
    意圖：價格偏低可能為買入好時機，偏高則風險增加。
    防呆：分母為 0 時設為 0。
    """
    # 3. 價格位置
    # -1. 總加分上限：0.05
    range_position = (target_price - low_price) / (high_price - low_price) if (high_price - low_price) > 0 else 0
    open_position = (target_price - open_price) / open_price if open_price > 0 else 0
    if range_position < 0.4 or open_position < -0.005:
        score += 0.05
        reasons.append(f"目標價格 {target_price:.0f} 低於開盤價 {open_price:.0f} 或接近低點 {low_price:.0f}，位置偏低")
    elif range_position > 0.6:
        score -= 0.05
        reasons.append(f"目標價格 {target_price:.0f} 接近高點 {high_price:.0f}，位置稍高")

    """
    邏輯：計算今日均價相對昨日收盤的變化率。
        close_factor > 0.005：上漲超過 0.5%，加 0.15。
        close_factor < -0.005：下跌超過 0.5%，減 0.05。
    中間範圍（-0.5% ~ 0.5%）：不調整。
    意圖：均價上漲反映趨勢向上，增強信心；下跌則減弱信心。
    防呆：昨日收盤為 0 時設為 0。
    """
    # 4. 昨日收盤影響
    # -1. 總加分上限：0.05
    close_factor = (avg_price - yesterday_close) / yesterday_close if yesterday_close > 0 else 0
    if close_factor > 0.005:
        score += 0.05
        reasons.append(f"今日均價 {avg_price:.2f} 較昨日收盤 {yesterday_close:.0f} 上漲 {close_factor:.2%}，正面影響")
    elif close_factor < -0.005:
        score -= 0.05
        reasons.append(f"今日均價 {avg_price:.2f} 較昨日收盤 {yesterday_close:.0f} 下跌 {close_factor:.2%}，些許負面影響")

    """
    邏輯：計算大盤外盤量佔比，類似個股內外盤。
        index_pressure_ratio > 0.55：大盤買盤強，加 0.1。
        index_pressure_ratio < 0.45：大盤賣壓重，減 0.05。
    意圖：大盤走勢影響個股，大盤強勢提升信心。
    調整：權重較低（0.1/0.05），因為大盤是間接因素。
    """
    # 5. 大盤影響
    # -1. 總加分上限：0.15
    """
    index_pressure_ratio = index_outer_vol / (index_inner_vol + index_outer_vol) if (index_inner_vol + index_outer_vol) > 0 else 0
    if index_pressure_ratio > 0.55:
        score += 0.15
        reasons.append(f"大盤外盤量 {index_outer_vol} 佔比 {index_pressure_ratio:.2%}，大盤買盤偏強")
    elif index_pressure_ratio < 0.45:
        score -= 0.15
        reasons.append(f"大盤內盤量 {index_inner_vol} 佔比 {1-index_pressure_ratio:.2%}，大盤賣壓稍重")
    """
    if index_current_price > index_open_price:
        score += 0.15
        reasons.append(f"大盤成交價 {index_current_price} 大於 大盤開盤價 {index_open_price}，大盤買盤偏強")
    else:
        score -= 0.15
        reasons.append(f"大盤成交價 {index_current_price} 小於 大盤開盤價 {index_open_price}，大盤買盤偏弱")

    """
    邏輯：
        base_suggested_price：以均價為基礎，根據訂單簿不平衡（imbalance）、昨日收盤變化（close_factor）、大盤壓力比（index_pressure_ratio）調整，每項乘以 0.01（即 1% 影響）。
        suggested_price = min(base_suggested_price, current_price)：取計算值與成交價較低者。
        math.floor：價格取整數（符合台股跳動單位）。
        限制範圍：不得低於當日最低價或高於最高價。
    意圖：綜合市場因素提供一個合理的建議買入價格。
    """
    # 6. 建議價格計算（調整權重）
    # base_suggested_price = avg_price * (1 + imbalance * 0.01 + close_factor * 0.01 + index_pressure_ratio * 0.01)
    base_suggested_price = avg_price * (1 + imbalance * 0.015 + close_factor * 0.015)
    suggested_price = min(base_suggested_price, current_price)
    suggested_price = math.floor(suggested_price)  # 整數跳動
    if suggested_price < low_price:
        suggested_price = low_price
    elif suggested_price > high_price:
        suggested_price = high_price

    """
    邏輯：
        price_diff：目標價格與建議價格的差距。
        price_diff <= 1：差距 ≤ 1 元，加 0.3（大幅提升信心）。
        price_diff <= 3：差距 ≤ 3 元，加 0.15。
    意圖：目標價格若接近建議價格，表示選擇合理，信心更高。
    """
    # 7. 與建議價格的接近程度
    # 1. 總加分上限：0
    price_diff = abs(target_price - suggested_price)
    if price_diff <= 1:
        # score += 0.3
        reasons.append(f"目標價格 {target_price:.0f} 與建議價格 {suggested_price:.0f} 差距僅 {price_diff:.0f}，極接近最佳買點，值得關注")
    elif price_diff <= 3:
        # score += 0.15
        reasons.append(f"目標價格 {target_price:.0f} 與建議價格 {suggested_price:.0f} 差距 {price_diff:.0f}，在合理範圍內，可考慮")
    
    score = max(0, min(1, score))
    
    return score, reasons, suggested_price

def submit_data(ticker='2330'):
    global window_alive, result_label, entries
    identify_name = 'submit_data_' + ticker
    logger = setup_logger(ticker, identify_name)
    if not window_alive:
        logger.warning("視窗已關閉，無法執行 submit_data_" + ticker)
        return
    try:
        # 修正順序與 default_values 一致
        yesterday_close = float(entries[0].get())# "昨日收盤價："
        current_price = float(entries[1].get())  # "成交價："
        avg_price = float(entries[2].get())      # "均價："
        inner_vol = float(entries[3].get())      # "內盤量："
        outer_vol = float(entries[4].get())      # "外盤量："
        opening_vol = float(entries[5].get())      # "開盤成交量："
        low_price = float(entries[6].get())      # "最低價："
        high_price = float(entries[7].get())     # "最高價："
        open_price = float(entries[8].get())     # "開盤價："
        bid_vol = float(entries[9].get())        # "委買量："
        ask_vol = float(entries[10].get())        # "委賣量："
        index_current_price = float(entries[11].get()) # "大盤成交價："
        index_inner_vol = float(entries[12].get()) # "大盤內盤量："
        index_outer_vol = float(entries[13].get()) # "大盤外盤量："
        index_low_price = float(entries[14].get()) # "大盤最低價："
        index_high_price = float(entries[15].get())# "大盤最高價："
        index_open_price = float(entries[16].get())    # "大盤開盤價："
        target_price = float(entries[17].get())  # "預期要買的價格："

        # 防呆檢查
        if target_price < low_price:
            raise ValueError(f"預期要買的價格 {target_price} 低於最低價 {low_price}")
        if low_price > open_price:
            raise ValueError(f"最低價 {low_price} 高於開盤價 {open_price}")
        if high_price < open_price:
            raise ValueError(f"最高價 {high_price} 低於開盤價 {open_price}")
        if high_price < low_price:
            raise ValueError(f"最高價 {high_price} 低於最低價 {low_price}")
        if target_price > current_price:
            raise ValueError(f"預期要買的價格 {target_price} 高於成交價 {current_price}，應小於等於成交價")

        # 假設的信心分數計算函數
        score, reasons, suggested_price = calculate_confidence_score(
            yesterday_close, avg_price, inner_vol, outer_vol, opening_vol, low_price, high_price, open_price, bid_vol, ask_vol, target_price,
            current_price, index_current_price, index_inner_vol, index_outer_vol, index_low_price, index_high_price, index_open_price
        )
        
        if suggested_price < low_price:
            suggested_price = low_price
            reasons.append(f"建議價格調整至最低價 {low_price:.0f}，因計算值過低")
        
        result_text = (f"您預期要買的價格為：{target_price:.0f}\n"
                       f"信心分數為：{score:.2f} (滿分 1.0)\n"
                       f"理由說明：\n" + "\n".join([f"- {r}" for r in reasons]) + "\n"
                       f"系統建議更加適當的價格為：{suggested_price:.0f}\n\n"
                       f"以上評估只供操作者自行參考，不代表任何投資價格的準則判斷")
        result_label.config(text=result_text)
        
        logger.info(f"輸入數據：均價={avg_price}, 內盤={inner_vol}, 外盤={outer_vol}, 開盤成交量={opening_vol} "
                    f"最低={low_price}, 最高={high_price}, 開盤={open_price}, 委買={bid_vol}, "
                    f"委賣={ask_vol}, 昨日收盤={yesterday_close}, 目標價={target_price}, "
                    f"成交價={current_price}, 大盤內盤={index_inner_vol}, 大盤外盤={index_outer_vol}, "
                    f"大盤最低={index_low_price}, 大盤最高={index_high_price}, 大盤當下={index_current_price}, "
                    f"大盤開盤={index_open_price}")
        logger.info(f"信心分數={score:.2f}, 建議價格={suggested_price:.0f}")
        
    except ValueError as e:
        messagebox.showerror("輸入錯誤", str(e))
    except tk.TclError:
        logger.error("控件已銷毀，無法獲取輸入數據")

def on_closing():
    global window_alive, entries
    window_alive = False
    entries = []
    root.destroy()

def refresh_data_wantgoo(ticker):
    global entries
    if not window_alive:
        return

    api_data_wantgoo = wantgoo_all_quote_info(ticker)
    if api_data_wantgoo:
        entries[11].delete(0, tk.END) # 大盤成交價
        entries[11].insert(0, int(api_data_wantgoo.get('close', 0)))
        entries[14].delete(0, tk.END) # 大盤最低價
        entries[14].insert(0, int(api_data_wantgoo.get('low', 0)))
        entries[15].delete(0, tk.END) # 大盤最高價
        entries[15].insert(0, int(api_data_wantgoo.get('high', 0)))
        entries[16].delete(0, tk.END) # 大盤開盤價
        entries[16].insert(0, int(api_data_wantgoo.get('open', 0)))
    else:
        messagebox.showwarning("刷新失敗", f"無法從 wantgoo API 獲取 {ticker} 的最新數據")

def refresh_data_twse(ticker='2330'):
    global entries
    if not window_alive:
        return

    api_data = twse_getStockInfo(ticker)
    if api_data and 'msgArray' in api_data and len(api_data['msgArray']) > 0:
        stock_info = api_data['msgArray'][0]
        # 更新相關欄位
        entries[0].delete(0, tk.END)  # 昨日收盤價
        entries[0].insert(0, stock_info.get('y', '0'))
        entries[1].delete(0, tk.END)  # 成交價
        entries[1].insert(0, stock_info.get('z', '0'))
        entries[6].delete(0, tk.END)  # 最低價
        entries[6].insert(0, stock_info.get('l', '0'))
        entries[7].delete(0, tk.END)  # 最高價
        entries[7].insert(0, stock_info.get('h', '0'))
        entries[8].delete(0, tk.END)  # 開盤價
        entries[8].insert(0, stock_info.get('o', '0'))
        entries[17].delete(0, tk.END) # 預期要買的價格
        entries[17].insert(0, stock_info.get('z', '0'))
    else:
        messagebox.showwarning("刷新失敗", f"無法從 TWSE API 獲取 {ticker} 的最新數據")

def refresh_data_fugle(ticker):
    global entries
    if not window_alive:
        return

    api_data = fugle_marketdata(ticker)
    if api_data:
        total_data = api_data.get('total', {})  # total 是字典，直接提取
        total_vol = int(total_data.get('tradeVolume', 0))  # 總量
        inner_vol = int(total_data.get('tradeVolumeAtBid', 0))  # 內盤量：賣方主動成交
        outer_vol = int(total_data.get('tradeVolumeAtAsk', 0))  # 外盤量：買方主動成交
        opening_vol = total_vol - (inner_vol + outer_vol)  # 開盤成交量
        entries[0].delete(0, tk.END)  # 昨日收盤價
        entries[0].insert(0, str(api_data.get('previousClose', '0')))
        entries[1].delete(0, tk.END)  # 成交價
        entries[1].insert(0, str(api_data.get('lastTrade', {}).get('price', '0')))
        entries[2].delete(0, tk.END)  # 均價
        entries[2].insert(0, str(api_data.get('avgPrice', '0')))
        entries[3].delete(0, tk.END)  # 內盤量
        entries[3].insert(0, inner_vol)
        entries[4].delete(0, tk.END)  # 外盤量
        entries[4].insert(0, outer_vol)
        entries[5].delete(0, tk.END)  # 開盤成交量
        entries[5].insert(0, opening_vol)
        entries[6].delete(0, tk.END)  # 最低價
        entries[6].insert(0, str(api_data.get('lowPrice', '0')))
        entries[7].delete(0, tk.END)  # 最高價
        entries[7].insert(0, str(api_data.get('highPrice', '0')))
        entries[8].delete(0, tk.END)  # 開盤價
        entries[8].insert(0, str(api_data.get('openPrice', '0')))
        entries[9].delete(0, tk.END)  # 委買量
        entries[9].insert(0, str(api_data.get('bids', [{}])[0].get('size', '0')))
        entries[10].delete(0, tk.END)  # 委賣量
        entries[10].insert(0, str(api_data.get('asks', [{}])[0].get('size', '0')))
        entries[17].delete(0, tk.END) # 預期要買的價格
        entries[17].insert(0, str(api_data.get('lastTrade', {}).get('price', '0')))
    else:
        messagebox.showwarning("刷新失敗", f"無法從 FUGLE API 獲取 {ticker} 的最新數據")

def clear_data():
    """清空數據：將所有欄位值清空"""
    global entries
    if not window_alive:
        return
    
    for entry in entries:
        entry.delete(0, tk.END)

def create_gui(ticker='2330'):
    global root, result_label, entries, window_alive
    
    root = tk.Tk()
    root.title(f"{ticker} 交易信心分數工具")
    
    default_values = {
        "昨日收盤價：": "0",
        "成交價：": "0",
        "均價：": "0",
        "內盤量：": "0",
        "外盤量：": "0",
        "開盤成交量：": "0",
        "最低價：": "0",
        "最高價：": "0",
        "開盤價：": "0",
        "委買量：": "0",
        "委賣量：": "0",
        "大盤成交價：": "0",
        "大盤內盤量：": "0",
        "大盤外盤量：": "0",
        "大盤最低價：": "0",
        "大盤最高價：": "0",
        "大盤開盤價：": "0",
        "預期要買的價格：": ""
    }

    """
    api_data = wantgoo_all_quote_info(ticker)
    time_param = str(api_data.get('time', '1741671000000')) if api_data else '1741671000000'
    topfive_data = historical_topfivepieces(ticker, time_param)
    
    if api_data:
    # if api_data and topfive_data:
        default_values = {
            "昨日收盤價：": str(api_data.get('previousClose', 0)),
            "成交價：": str(api_data.get('close', 0)),
            "均價：": str(api_data.get('close', 0)),
            "內盤量：": "0",
            "外盤量：": "0",
            "開盤成交量：": "0",
            "最低價：": str(api_data.get('low', 0)),
            "最高價：": str(api_data.get('high', 0)),
            "開盤價：": str(api_data.get('open', 0)),
            "委買量：": "0",
            "委賣量：": "0",
            "大盤開盤價：": "0",
            "大盤內盤量：": "0",
            "大盤外盤量：": "0",
            "大盤最低價：": "0",
            "大盤最高價：": "0",
            "大盤成交價：": "0",
            "預期要買的價格：": str(api_data.get('close', 0))
        }
    """

    """
    api_data = twse_getStockInfo(ticker)
    
    # 若 API 抓取成功，解析所需欄位
    if api_data and 'msgArray' in api_data and len(api_data['msgArray']) > 0:
        stock_info = api_data['msgArray'][0]  # 取第一筆數據
        default_values = {
            "昨日收盤價：": stock_info.get('y', '0'),      # 昨日收盤價
            "成交價：": stock_info.get('z', '0'),          # 成交價
            "均價：": "0",                                 # TWSE無均價，暫設為0
            "內盤量：": "0",                               # TWSE無內外盤，暫設為0
            "外盤量：": "0",
            "開盤成交量：": "0",
            "最低價：": stock_info.get('l', '0'),          # 最低價
            "最高價：": stock_info.get('h', '0'),          # 最高價
            "開盤價：": stock_info.get('o', '0'),          # 開盤價
            "委買量：": "0",                               # TWSE無五檔，暫設為0
            "委賣量：": "0",
            "大盤成交價：": "0",
            "大盤內盤量：": "0",                           # TWSE無大盤數據，暫設為0
            "大盤外盤量：": "0",
            "大盤最低價：": "0",
            "大盤最高價：": "0",
            "大盤開盤價：": "0",
            "預期要買的價格：": stock_info.get('z', '0')   # 預設用成交價
        }
        root.title(f"{ticker} {stock_info.get('n', '')} 交易信心分數工具")
    """
    
    # """
    api_data = fugle_marketdata(ticker)
    api_data_wantgoo = wantgoo_all_quote_info('WTX&')

    # 若 API 抓取成功，解析所需欄位
    if api_data:
        total_data = api_data.get('total', {})  # total 是字典，直接提取
        total_vol = int(total_data.get('tradeVolume', 0))  # 總量
        inner_vol = int(total_data.get('tradeVolumeAtBid', 0))  # 內盤量：賣方主動成交
        outer_vol = int(total_data.get('tradeVolumeAtAsk', 0))  # 外盤量：買方主動成交
        opening_vol = total_vol - (inner_vol + outer_vol)  # 開盤成交量
        default_values = {
            "昨日收盤價：": str(api_data.get('previousClose', '0')),
            "成交價：": str(api_data.get('lastTrade', {}).get('price', '0')),
            "均價：": str(api_data.get('avgPrice', '0')),
            "內盤量：": inner_vol,
            "外盤量：": outer_vol,
            "開盤成交量：": opening_vol,
            "最低價：": str(api_data.get('lowPrice', '0')),
            "最高價：": str(api_data.get('highPrice', '0')),
            "開盤價：": str(api_data.get('openPrice', '0')),
            "委買量：": str(api_data.get('bids', [{}])[0].get('size', '0')),
            "委賣量：": str(api_data.get('asks', [{}])[0].get('size', '0')),
            "大盤成交價：": int(api_data_wantgoo.get('close', 0)),
            "大盤內盤量：": "0",
            "大盤外盤量：": "0",
            "大盤最低價：": int(api_data_wantgoo.get('low', 0)),
            "大盤最高價：": int(api_data_wantgoo.get('high', 0)),
            "大盤開盤價：": int(api_data_wantgoo.get('open', 0)),
            "預期要買的價格：": int(api_data.get('lastTrade', {}).get('price', '0'))
        }
        root.title(f"{ticker} {str(api_data.get('name', ''))} 交易信心分數工具")
    # """

    labels = list(default_values.keys())
    entries = []

    for i, label_text in enumerate(labels):
        tk.Label(root, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(root)
        entry.insert(0, default_values[label_text])
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries.append(entry)

    # 注意：這裡不再提前分配變數，而是讓 submit_data 直接使用 entries 索引

    tk.Button(root, text="進行評估", command=lambda: submit_data(ticker)).grid(row=len(labels), column=0, padx=5, pady=10)
    tk.Button(root, text="刷新數據", command=lambda: [refresh_data_fugle(ticker), refresh_data_wantgoo('WTX&')]).grid(row=len(labels), column=1, padx=5, pady=10)
    tk.Button(root, text="清空數據", command=clear_data).grid(row=len(labels), column=2, padx=5, pady=10)

    result_label = tk.Label(root, text="", justify="left", wraplength=400)
    result_label.grid(row=len(labels)+1, column=0, columnspan=3, padx=5, pady=5)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

def twse_getStockInfo(ticker='2330'):
    """
    從 twse API 抓取指定 ticker 的股票相關資訊。
    
    參數:
        ticker (str): 股票代碼，預設為 '2330'。
    
    回傳跟欄位範例說明:
        dict: 符合指定 ticker 的股票資訊，若未找到則回傳 None。
        {
          "msgArray": [
            {
              "@": "2330.tw",
              "c": "2330",
              "h": "995.0000", # 最高價
              "l": "976.0000", # 最低價
              "n": "台積電",
              "o": "980.0000", # 開盤價
              "nf": "台灣積體電路製造股份有限公司",
              "y": "971.0000", # 昨天收盤價
              "z": "988.0000" # 成交價
              ...略...
            }
          ],
          ...略...
        }

    例外:
        拋出 requests.exceptions.RequestException 如果網路請求失敗。
    """

    # 檔名識別
    identify_name = 'twse_getStockInfo'
    
    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_" + ticker + ".tw"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        # 發送請求並確保請求成功
        response = requests.get(url, headers=headers, timeout=10)  # 發送請求
        response.raise_for_status()  # 確保請求成功
        
        # 解析 JSON 回應
        data = response.json()
        logger.info(f"成功抓取 '{ticker}' 的數據：{data}")
        return data
    
    except requests.RequestException as e:
        logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        logger.error(f"數據解析錯誤：{str(e)}")
        return None

def fugle_marketdata(ticker='2330'):
    """
    從 fugle API 抓取指定 ticker 的股票相關資訊。
    
    參數:
        ticker (str): 股票代碼，預設為 '2330'。
    
    回傳跟欄位範例說明:
        dict: 符合指定 ticker 的股票資訊，若未找到則回傳 None。
        {
            'date': '2025-03-12',
            'symbol': '2330',
            'name': '台積電',
            'openPrice': 980,
            'highPrice': 995,
            'lowPrice': 976,
            'closePrice': 988,
            'avgPrice': 985.05,
            'previousClose': 971,
            'bids': [{'price': 987, 'size': 1}, {'price': 986, 'size': 6}, ...],  # 五檔委買
            'asks': [{'price': 988, 'size': 149}, {'price': 989, 'size': 169}, ...],  # 五檔委賣
            'total': {'tradeValue': 33528251000, 'tradeVolume': 34037, ...},
            'lastTrade': {'price': 988, 'size': 5128, 'time': 1741757400000000},
            'isClose': True,
            ...
        }

    例外:
        拋出 requests.exceptions.RequestException 如果網路請求失敗。
        
    參考網址：https://developer.fugle.tw/docs/key
    參考網址：https://developer.fugle.tw/docs/data/intro
    參考網址：https://github.com/fugle-dev/fugle-marketdata-python
    """
    
    # 檔名識別
    identify_name = 'fugle_marketdata'
    
    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 發送請求
    try:
        client = RestClient(api_key = os.getenv('FUGLE_API_TOKEN'))
        stock = client.stock  # Stock REST API client
        data = stock.intraday.quote(symbol=ticker)
        logger.info(f"成功抓取 '{ticker}' 的數據：{data}")
        return data
    except requests.RequestException as e:
        logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        logger.error(f"數據解析錯誤：{str(e)}")
        return None

def wantgoo_all_quote_info(ticker='2330'):
    """
    從 wantgoo API 抓取指定 ticker 的股票相關資訊。
    
    參數:
        ticker (str): 股票代碼，預設為 '2330'。
    
    回傳跟欄位範例說明:
        dict: 符合指定 ticker 的股票資訊，若未找到則回傳 None。
        {
          "id": "2330", # 股票代號
          "tradeDate": 1741622400000,
          "time": 1741671000000, # 時間參數
          "flat": 998.0, # 昨天收盤價
          "floor": 899.0,
          "ceil": 1095.0,
          "open": 969.0, # 今天開盤價
          "high": 979.0, # 當下最高價
          "low": 963.0, # 當下最低價
          "close": 971.0, # 當下成交價
          "volume": 50047, # 總量
          "millionAmount": 48579.34,
          "previousClose": 998.0, # 昨天收盤價
          "previousVolume": 46774,
          "previousMillionAmount": 46661.0
        }

    例外:
        拋出 requests.exceptions.RequestException 如果網路請求失敗。
    參考網址：https://www.wantgoo.com/server-time
    """

    # 檔名識別
    identify_name = 'wantgoo_all_quote_info'
    
    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    url = "https://www.wantgoo.com/investrue/all-quote-info"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        # 發送請求並確保請求成功
        response = requests.get(url, headers=headers, timeout=10)  # 發送請求
        response.raise_for_status()  # 確保請求成功
        
        # 解析 JSON 回應
        data = response.json()
        
        # 在 JSON 列表中尋找符合 ticker 的資料
        for item in data:
            if item.get('id') == ticker:
                logger.info(f"成功抓取 '{ticker}' 的數據：{item}")
                return item

        # 若未找到指定 ticker，回傳 None
        logger.error(f"未找到 id為 '{ticker}' 的資料")
        return None
    
    except requests.RequestException as e:
        logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        logger.error(f"數據解析錯誤：{str(e)}")
        return None

def historical_topfivepieces(ticker='2330', time='1741671000000'):
    """
    從 wantgoo API 抓取指定 ticker 的股票相關資訊。（上下五檔）
    
    參數:
        ticker (str): 股票代碼，預設為 '2330'。
    
    回傳跟欄位範例說明:
        dict: 符合指定 ticker 的股票資訊，若未找到則回傳 None。
        {
          "bidPrice1": 970, # 下1檔價格
          "bidVolume1": 1720, # 下1檔委買量
          "bidPrice2": 969, # 下2檔價格
          "bidVolume2": 397, # 下2檔委買量
          "bidPrice3": 968, # 下3檔價格
          "bidVolume3": 990, # 下3檔委買量
          "bidPrice4": 967, # 下4檔價格
          "bidVolume4": 381, # 下4檔委買量
          "bidPrice5": 966,, # 下5檔價格
          "bidVolume5": 719 # 下5檔委買量
          "askPrice1": 971, # 上1檔價格
          "askVolume1": 48, # 上1檔委買量
          "askPrice2": 972, # 上2檔價格
          "askVolume2": 73, # 上2檔委買量
          "askPrice3": 973, # 上3檔價格
          "askVolume3": 131, # 上3檔委買量
          "askPrice4": 974, # 上4檔價格
          "askVolume4": 141, # 上4檔委買量
          "askPrice5": 975, # 上5檔價格
          "askVolume5": 1421, # 上5檔委買量
          "dealOnBidPrice": 26965,
          "dealOnAskPrice": 22972
        }

    例外:
        拋出 requests.exceptions.RequestException 如果網路請求失敗。
    """

    # 檔名識別
    identify_name = 'historical_topfivepieces'
    
    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    url = "https://www.wantgoo.com/investrue/" + ticker + "/historical-topfivepieces?v=" + time
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        # 發送請求並確保請求成功
        response = requests.get(url, headers=headers, timeout=10)  # 發送請求
        response.raise_for_status()  # 確保請求成功
        
        # 解析 JSON 回應
        data = response.json()
        
        logger.info(f"成功抓取 '{ticker}' 的數據：{data}")
        return data

        # 若未找到指定 ticker，回傳 None
        logger.error(f"未找到 id為 '{ticker}' 的資料")
        return None
    
    except requests.RequestException as e:
        logger.error(f"網路請求失敗：{str(e)}")
        return None
    except (ValueError, AttributeError) as e:
        logger.error(f"數據解析錯誤：{str(e)}")
        return None

def get_average_price(ticker='2330'):
    session = requests.Session()
    url = f"https://www.wantgoo.com/investrue/{ticker}/average-price"
    main_page = f"https://www.wantgoo.com/stock/{ticker}"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
        'Referer': main_page,
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'X-Requested-With': 'XMLHttpRequest'
    }

    try:
        # 先訪問主頁獲取 cookie
        session.get(main_page, headers=headers, timeout=10)
        # 發送 API 請求
        response = session.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        data = response.json()
        avg_price = data.get('value', None)
        print(f"抓取到的均價: {avg_price}")
        return avg_price
    except requests.RequestException as e:
        print(f"抓取失敗: {e}")
        return None

def get_average_price_selenium(ticker='2330'):
    url = f"https://www.wantgoo.com/stock/{ticker}"
    service = Service(ChromeDriverManager().install())
    
    options = Options()
    options.add_argument('--headless')  # 無頭模式
    options.add_argument('--disable-gpu')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/134.0.0.0 Safari/537.36')
    options.add_argument('--disable-blink-features=AutomationControlled')  # 隱藏自動化特徵
    
    driver = webdriver.Chrome(service=service, options=options)
    
    try:
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
        })  # 隱藏 webdriver 屬性
        driver.get(url)
        driver.implicitly_wait(10)
        avg_price_elem = driver.find_element(By.ID, 'averagePrice')
        avg_price = avg_price_elem.text
        print(f"抓取到的均價: {avg_price}")
        return avg_price
    except Exception as e:
        print(f"抓取失敗: {e}")
        return None
    finally:
        driver.quit()

def call_chatgpt_api(prompt):
    """
    參考網址：https://platform.openai.com/settings/organization/api-keys
    """
    """調用 ChatGPT API 並返回分析"""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "gpt-4o",  # 使用 GPT-4o，若成本考量可改成 gpt-3.5-turbo
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 2000
    }
    response = requests.post(API_ENDPOINT, headers=headers, json=payload)
    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]
    else:
        print(f"API 錯誤: {response.status_code} - {response.text}")
        # return f"API 錯誤: {response.status_code} - {response.text}"

def data_report(ticker='2330'):
    # 檔名識別
    identify_name = 'data_report'

    # 關聯log行為
    logger = setup_logger(ticker, identify_name)

    # 關聯字型行為
    setup_chinese_font(logger)

    # 指定要合併的內容：純文字標籤或 CSV 檔案路徑
    csv_files = [
        # {"label": "重大歷史事件：\n"},
        # {"label": "輝達17日宣布投資英特爾50億美元"},
        # {"label": "輝達投資1000億美元！攜手OpenAI打造AI超級資料中心\n"},
        {"label": "以下是 台積電 相關的歷史數據紀錄\n"},
        {"label": "目前買入均價為：0"},
        {"label": "目前買入張數為：0"},
        {"label": "剩餘加碼張數為：5（3張現股+2張融資）\n"},
        {"label": "最近交易紀錄：\n"},
        {"label": "[台股台積電歷史股價]\n"},
        {"path": r"output\2330\2330_yfinance_daily.csv"},
        {"label": "\n[台指期盤後]\n"},
        {"path": r"output\WTXP&\WTXP&.csv"},
        {"label": "\n[台積電期貨盤後]\n"},
        {"path": r"output\WCDFP&\WCDFP&.csv"},
        {"label": "\n[美股台積電ADR歷史股價]\n"},
        {"path": r"output\TSM\TSM_yfinance_daily.csv"},
        {"label": "\n[NASDAQ指數]\n"},
        {"path": r"output\^IXIC\^IXIC_yfinance_daily.csv"},
        {"label": "\n[費城半導體指數]\n"},
        {"path": r"output\^SOX\^SOX_yfinance_daily.csv"},
        {"label": "\n[美金兌換台幣匯率]\n"},
        {"path": r"output\exchange_rate\USD\USD_exchange_rate.csv"},
        {"label": "\n[台股外資空單口數]\n"},
        {"path": r"output\futures\foreign_empty_orders.csv"},
        {"label": "\n[融資券數]\n"},
        {"path": r"output\2330\2330_finmind_margin_purchase.csv"},
        {"label": "\n[法人買賣超]\n"},
        {"path": r"output\2330\2330_finmind_institutional.csv"},
        {"label": "\n[kd 9]\n"},
        {"path": r"output\2330\2330_yfinance_kd_9.csv"},
        {"label": "\n[kd 14]\n"},
        {"path": r"output\2330\2330_yfinance_kd_14.csv"},
        {"label": "\n[ma]\n"},
        {"path": r"output\2330\2330_yfinance_ma_5_200.csv"},
        {"label": "\n[rsi 9]\n"},
        {"path": r"output\2330\2330_yfinance_rsi_9.csv"},
        {"label": "\n[rsi 14]\n"},
        {"path": r"output\2330\2330_yfinance_rsi_14.csv"},
        {"label": "\n[macd]\n"},
        {"path": r"output\2330\2330_yfinance_macd_12_26_9.csv"},
        {"label": "\n[bollinger]\n"},
        {"path": r"output\2330\2330_yfinance_bollinger_20_2.0.csv"},
        # {"label": "\n[obv]\n"},
        # {"path": r"output\2330\2330_yfinance_obv.csv"},
        {"label": "\n請根據上數據來分析 台積電 接下來的可能走勢跟因應狀況，配合當下時空背景跟國際大事等因素做出一些建議跟看法，特別是接下來開盤時的價格區間可能的樂觀悲觀中性百分比為何"},
    ]

    # 設定輸出檔案路徑
    output_file = r"台積電2330數據報告.txt"

    # 確認檔案是否存在（僅檢查 path 項目）
    existing_files = [item for item in csv_files if item.get("path") and os.path.exists(item["path"])]
    if not any("path" in item for item in csv_files):
        logger.warning(f"沒有找到 {ticker} 所指定的 CSV 檔案的數據")
    else:
        # 打開輸出檔案並寫入
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for i, item in enumerate(csv_files):
                if "label" in item:
                    # 寫入純文字標籤
                    outfile.write(f"{item['label']}\n")
                elif "path" in item and os.path.exists(item["path"]):
                    # 讀取並寫入 CSV 檔案內容
                    with open(item["path"], 'r', encoding='utf-8') as infile:
                        content = infile.read()
                        outfile.write(content)
                # 如果不是最後一個項目，添加換行符
                # if i < len(csv_files) - 1:
                    # outfile.write('\n')
        
        logger.info(f"已成功合併 {len(existing_files)} 個 CSV 檔案到 {output_file}")

# 測試區塊，以方便直接運行 判斷目前這個 .py 檔案是不是直接被執行
if __name__ == "__main__":
    process_stock_data()
    night_trading('WTXP&')  # 台指期盤後
    night_trading('WCDFP&')  # 台積電期貨盤後
    process_stock_data('^TWII')  # 加權指數
    process_stock_data('TSM')  # 台積電ADR
    # process_stock_data('^DJI')  # 道瓊工業指數
    # process_stock_data('YM=F')  # 道瓊期貨指數
    # process_stock_data('^GSPC')  # S&P 500
    # process_stock_data('ES=F')  # S&P 500期貨指數
    process_stock_data('^IXIC')  # NASDAQ指數
    process_stock_data('NQ=F')  # NASDAQ期貨指數
    process_stock_data('^SOX')  # 費城半導體指數
    exchange_rate()  # 台幣兌換美金歷史匯率
    number_of_empty_orders()  # 外資每日空單口數
    taiwan_stock_daily()
    taiwan_stock_margin_purchase_short_sale()
    taiwan_stock_institutional_investors()
    # rsv(9)
    # rsv(14)
    kd(9)
    kd(14)
    ma()
    rsi(9)
    rsi(14)
    macd()
    bollinger()
    obv()
    data_report()