import os
import shutil
import logging
from datetime import datetime, timedelta
import sys

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


def clean_old_logs(log_dir='logs', days_to_keep=30):
    """
    清理超過指定天數的日誌資料夾。
    
    參數：
        log_dir (str): 日誌根目錄，預設為 'logs'
        days_to_keep (int): 保留天數，預設為 30 天
    """
    if not os.path.exists(log_dir):
        logger.info(f"日誌目錄 {log_dir} 不存在，無需清理")
        return

    for folder in os.listdir(log_dir):
        folder_path = os.path.join(log_dir, folder)
        try:
            folder_date = datetime.strptime(folder, '%Y%m%d')
            if datetime.now() - folder_date > timedelta(days=days_to_keep):
                shutil.rmtree(folder_path)
                logger.info(f"已刪除過期日誌資料夾：{folder_path}")
            else:
                logger.info(f"保留日誌資料夾：{folder_path}")
        except ValueError:
            logger.warning(f"跳過無效資料夾：{folder}")
        except Exception as e:
            logger.error(f"刪除 {folder_path} 時發生錯誤：{e}")

if __name__ == "__main__":
    days = int(sys.argv[1]) if len(sys.argv) > 1 else 30
    clean_old_logs(days_to_keep=days) # python clean.py 60（保留 60 天）