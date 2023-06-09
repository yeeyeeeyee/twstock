import twstock

stock_data = twstock.realtime.get(["0050", "0052","0055","0056","2330","006204","006208",
                                   "00637L","00646","00679B","00690","00692","00701",
                                   "00708L",'00712','00713','00728','00730',"00731",])

#print(stock_data["0050"]['realtime'])


for stock_code in stock_data:
    #最後一個會是succcess,所以在最後一個前停下來
    if "success" == stock_code:
        break

    best_ask_volume = stock_data[stock_code]['info']["name"]
    print(f"{stock_code} 的 best_ask_volume 為: {best_ask_volume}")
    


















'00733',
'00751B',
'00752',
'00757',
'00770',
'00773B',
'00830',
"00850"
'00878'
'00881'
'00882'
'00885'
'00888'
'00892'
'00894'
'00900'
'00907'
'00771'
