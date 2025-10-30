# 載入必要的套件
import time,requests,os,datetime,random
from json import loads
import pandas as pd

c_datetime=datetime.datetime.now()
c_day=c_datetime.strftime('%Y%m%d')

first_write=True

while c_day>'20250304':
    
    c_datetime -= datetime.timedelta(1)
    c_day=c_datetime.strftime('%Y%m%d')

    # 設置標頭，模擬瀏覽器
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    print(c_day)
    url='https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date='+c_day+'&selectType=ALL&_=1653293140736'
    
    html=requests.get(url, headers=headers)
    jsdata=loads(html.text)
    
    if 'data' not in jsdata.keys():
        print(c_day,'爬蟲失敗')
        continue            
    
    data=pd.DataFrame(jsdata['data'],columns=jsdata['fields'])
    data['日期']=c_day
    
    if first_write:
        data.to_csv('融資融券爬蟲資料.csv',encoding='cp950',index=False) 
        first_write=False
    else:
        data.to_csv('融資融券爬蟲資料.csv',encoding='cp950',mode='a',index=False,header=False) 

    time.sleep(random.randint(5,10))
    
