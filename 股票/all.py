#--------套件 getstock---------#
import openpyxl
import pandas as pd
import twstock
def get(name,search_date):
    print(twstock.codes[name])
    print("----------------------")
    stock_0050 = twstock.Stock(name)

    transaction= stock_0050.transaction[search_date]
    date=stock_0050.date[search_date]
    open0050=stock_0050.open[search_date]
    hight=stock_0050.high[search_date]
    low=stock_0050.low[search_date]
    close=stock_0050.close[search_date]
    turnover=stock_0050.turnover[search_date]
    change=stock_0050.change[search_date]
    

    print(f'transaction:{transaction}\
------date:{date}-----open:{open0050}----hight:{hight}--low:{low}\n \
-close:{close}------turnover:{turnover}---change:{change}-----振幅:{round(((hight-low)/low)*100,3)}')

code=input("input code:")
date=input("search:")
try:
    get(code,int(date))
    
except TimeoutError:
    print("timeout Error")