import twstock

class get_stock_ed():
    def __init__(self,code,during_star=-1,during_end=None) -> None:
        self.__code=code
        self.__during_star=during_star
        self.__during_end=during_end
        self.__stock=twstock.Stock(self.__code)

#--------------------------------------   
#           收盤後資訊
 
    #獲得輸入股票基本資訊
    def info(self)->list:
        return twstock.codes[self.__code]
    #股票id
    def id(self)->str:
        return self.__stock.sid
    #獲取日期
    def date(self)->list:
        return self.__stock.date[self.__during_star:self.__during_end]
    #獲得漲跌
    def change(self):
        return self.__stock.change[self.__during_star:self.__during_end]
    #獲得高點
    def high(self)->list:
        return self.__stock.high[self.__during_star:self.__during_end]
    #獲得低點
    def low(self)->list:
        return self.__stock.low[self.__during_star:self.__during_end]
    #獲得成交
    def price(self)->list:
        return self.__stock.price[self.__during_star:self.__during_end]
    #獲得開盤
    def open(self)->list:
        return self.__stock.open[self.__during_star:self.__during_end]
    #獲得收盤
    def close(self)->list:
        return self.__stock.close[self.__during_star:self.__during_end]

    #獲得振幅%
    # 振幅=((高點-低點)/低點)*100 ,再以振幅四捨五入小數第3位數
    def amplitude(self)->list:
        all_amplitude=[]
        for i in range(len(get_stock_ed.high(self))):
            amplitude= round(((get_stock_ed.high(self)[i]-get_stock_ed.low(self)[i])/get_stock_ed.low(self)[i])*100,3)
            all_amplitude.append(amplitude)
        return all_amplitude
#               收盤後資訊
#-------------------------------------
   
class get_stock_ing():
    def __init__(self,code) -> None:
        self.code=code
    def get(self):
        if type(self.code)==list:
            list_stock=[]
            for i in self.code:
                list_stock.append(twstock.realtime.get(i))
            return list_stock[0]
            
        else:
            return twstock.realtime.get(self.code)

        

    



stock=get_stock_ing("0050")
print(stock.get())

