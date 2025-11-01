import twstock
# 以台積電的股票代號建立 Stock 物件
stock = twstock.Stock('2330')  
print(stock.price)