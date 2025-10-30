import function
import sys

# 日記記錄
import logging

# 時間格式
from datetime import datetime

import os

# 動態設置日誌
current_date = datetime.now().strftime('%Y%m%d')
log_dir = f'logs/{current_date}'
os.makedirs(log_dir, exist_ok=True)
file_name = os.path.splitext(os.path.basename(__file__))[0]  # 獲取主檔名，例如 'stock'
log_file = f'{log_dir}/{file_name}.log'  # 例如 'logs/20250227/stock.log'

# 設置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),  # 指定 UTF-8 編碼
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

import inspect

def route_command():
    if len(sys.argv) < 2:
        logger.error("使用方式：python stock.py <function_name> [args...]")
        logger.error("範例：python stock.py process_stock_data 0052 2025-02-17 2025-02-27")
        sys.exit(1)

    # 獲取指令列參數
    func_name = sys.argv[1]  # 第一個參數是函數名稱
    args = sys.argv[2:]      # 後續參數是函數的輸入

    # 動態獲取函數
    try:
        func = getattr(function, func_name)
    except AttributeError:
        available_funcs = [name for name in dir(function) if inspect.isfunction(getattr(function, name)) and not name.startswith('_')]
        logger.error(f"錯誤：函數 '{func_name}' 在 function.py 中不存在")
        logger.error(f"可用函數：{available_funcs}")
        sys.exit(1)

    # 檢查是否為可調用對象
    if not callable(func):
        logger.warning(f"錯誤：'{func_name}' 不是一個可調用的函數")
        sys.exit(1)

    # 執行函數，傳入剩餘參數
    try:
        func(*args)
        logger.info(f"成功執行函數 '{func_name}'")
    except TypeError as e:
        logger.error(f"錯誤：參數錯誤 - {e}")
        logger.error(f"請確認 '{func_name}' 所需的參數數量和格式")
        sys.exit(1)

if __name__ == "__main__":
    route_command()