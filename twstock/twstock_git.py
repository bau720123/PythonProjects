from twstock import Stock
from twstock import BestFourPoint

stock = Stock('2330') # 擷取台積電股價
ma_p = stock.moving_average(stock.price, 5)
ma_c = stock.moving_average(stock.capacity, 5)
ma_p_cont = stock.continuous(ma_p)
ma_br = stock.ma_bias_ratio(5, 10)

print('計算五日均價：', ma_p)
print('計算五日均量：', ma_c)
print('計算五日均價持續天數：', ma_p_cont)
print('計算五日、十日乖離值：', ma_br)

period = stock.fetch_from(2025, 5)
print('2025年5月以來：', period)

bfp = BestFourPoint(stock)

fourTopBuyPoint = bfp.best_four_point_to_buy()
fourTopSellPoint = bfp.best_four_point_to_sell()
analyze= bfp.best_four_point()

print('判斷是否為四大買點：', fourTopBuyPoint)
print('判斷是否為四大賣點：', fourTopSellPoint)
print('綜合判斷：', analyze)

import twstock

realtime = twstock.realtime.get('2330')    # 擷取當前台積電股票資訊
# twstock.realtime.get(['2330', '2337', '2409'])  # 擷取當前三檔資訊

print('擷取當前台積電股票資訊：', realtime)