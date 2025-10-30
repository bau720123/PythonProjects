# 偵測系統平台
import platform

# 載入 繪圖 套件
import matplotlib.pyplot as plt

# 日記記錄
import logging

# 用來檢查檔案是否存在
import os

# 時間格式
from datetime import datetime, timedelta

# 比較運算符
import operator

# 定義運算符映射
COMPARE_OPERATORS = {
    '>': operator.gt,   # 大於
    '>=': operator.ge,  # 大於等於
    # '<': operator.lt,   # 小於
    # '<=': operator.le,  # 小於等於
    # '==': operator.eq,  # 等於
}

def setup_logger(ticker, identify_name):
    # 動態設置日誌
    current_date = datetime.now().strftime('%Y%m%d')
    log_dir = f'logs/{current_date}'
    os.makedirs(log_dir, exist_ok=True)
    log_file = f'{log_dir}/{ticker}_{identify_name}.log'

    # 配置 logger
    logger = logging.getLogger(f"{ticker}_{identify_name}")  # 每個 ticker 有獨立的 logger
    logger.setLevel(logging.DEBUG)
    if os.getenv('environment') == 'production':
        logger.setLevel(logging.INFO)
    logger.propagate = False  # 禁用傳播
    if not logger.handlers:  # 避免重複添加 handler
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler = logging.FileHandler(log_file, encoding='utf-8')  # 指定 UTF-8 編碼
        file_handler.setFormatter(formatter)
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)
    return logger
        
def setup_check_realtime_format(logger, start_date, end_date, comparison_operator='>'):
    """
    檢查日期格式是否正確，並根據 comparison_operator 參數比較 start_date 和 end_date。
    參數：
        logger: 日誌記錄器 (預設 None)
        start_date: 開始日期 (YYYY-MM-DD)
        end_date: 結束日期 (YYYY-MM-DD)
        comparison_operator: 比較方式 (預設 '>'，目前僅支援 '>', '>=')
    返回：成功則返回 True，失敗則記錄錯誤並返回 False。
    """
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError as e:
        err_msg = f"日期格式錯誤，應為 YYYY-MM-DD，收到 start_date={start_date}, end_date={end_date}"
        logger.error(err_msg)
        raise ValueError(err_msg) from e  # 保留原始異常原因

    # 動態選擇比較運算符
    if comparison_operator not in COMPARE_OPERATORS:
        err_msg = f"無效的比較運算符 {comparison_operator}，目前僅支援 '>', '>='"
        logger.error(err_msg)
        raise ValueError(err_msg)
    
    compare_func = COMPARE_OPERATORS[comparison_operator]
    if compare_func(start_dt, end_dt):
        err_msg = f"開始日期 {start_date} {comparison_operator} 結束日期 {end_date} 不符合條件"
        logger.error(err_msg)
        raise ValueError(err_msg)
    
    return True

def setup_folders(ticker):
    # 加入資料夾管理
    output_dir = f'output/{ticker}'
    os.makedirs(output_dir, exist_ok=True)  # 創建資料夾，若已存在則忽略

    return output_dir

def setup_chinese_font(logger):
    # 設置支援中文的字型
    system = platform.system()
    if system == 'Windows':
        plt.rcParams['font.family'] = 'Microsoft YaHei'
    elif system == 'Darwin':
        plt.rcParams['font.family'] = 'AppleGothic'
    elif system == 'Linux':
        plt.rcParams['font.family'] = 'Noto Sans CJK SC'
    else:
        plt.rcParams['font.family'] = 'sans-serif'
        if logger:
            logger.warning("未設置適配的中文字型，可能影響顯示")