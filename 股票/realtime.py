import twstock

stock_data = twstock.realtime.get(["0050", "0052","0056","2330"])

#print(stock_data["0050"]['realtime'])


for stock_code in stock_data:
    #最後一個會是succcess,所以在最後一個前停下來
    if "success" == stock_code:
        break

    best_ask_volume = stock_data[stock_code]['info']
    print(f"{stock_code} 的 best_ask_volume 為: {best_ask_volume}")
    
