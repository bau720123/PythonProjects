import twstock
# 以台積電的股票代號建立 Stock 物件
stock = twstock.Stock('2330')  
# 取得 2025 年 06 月的資料
stocklist = stock.fetch(2025,6)   
for s in stocklist:
    print(s.date.strftime('%Y-%m-%d'), end='\t')
    print(s.open, end='\t')
    print(s.high, end='\t')
    print(s.low, end='\t')
    print(s.close)