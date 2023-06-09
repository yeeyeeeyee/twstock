import twstock
import pandas as pd
stock=twstock.realtime.get(["0050","0052"])
#stock = twstock.realtime.get(['52', '55', '56', '6204', '6208', '00637L', '646', '00679B', '690', '692', '701', '00708L', '712', '713', '728', '730', '731', '733', '00751B', '752', '757', '770', '00772B', '00773B', '830', '850', '878', '881', '882', '885', '888', '892', '894', '900', '907', '771'])
#print(stock["0052"]["realtime"])
#print(stock)
for stock_code in ['0050', '0052']:
    best_ask_volume = stock[stock_code]['realtime']
    print(f"{stock_code} 的 best_ask_volume 為: {best_ask_volume}")

try:
    data = pd.DataFrame(stock["realtime"])
    #data.to_excel("data.xlsx", index=False)
except:
    print("error")
