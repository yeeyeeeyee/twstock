import twstock
#測試有沒有被擋住
print (twstock.realtime.get('0050'))

#測試歷史資料的日期
""" stock=twstock.Stock('0050')
print(stock.date[-1]) """