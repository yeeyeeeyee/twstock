import twstock
import time
import xlwings as xw

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

   
class get_stock_ing():
    def __init__(self,code) -> None:
        self.__code=code
    #------------------------------
    #       獲取盤中資料
    def get_all_information(self):
        return self.__code
    
    #獲得info裡面個別資料
    def get_info(self):
        return self.__code["info"]
    #獲得時間
    def get_time(self):
        time=self.__code["info"]["time"]
        time=time.split(" ")
        #回傳格式  ('2023-06-14', '14:30:00')
        return time[0],time[1]
    #獲得代號
    def get_code(self):
        return self.__code["info"]["code"]
    #獲得名稱
    def get_name(self):
        return self.__code["info"]["name"]
    
    #獲得realtime裡面個別資料
    def get_realtime(self):
        return self.__code["realtime"]
   
    #成交價
    def get_latest_trade_price(self):
        return self.__code["realtime"]["latest_trade_price"]
    #成交量
    def get_trade_volume(self):
        return self.__code["realtime"]["trade_volume"]
    
    #總成交量
    def get_accumulate_trade_volume(self):
        return self.__code["realtime"]["accumulate_trade_volume"]
    
    #買進價格
    def get_best_bid_price(self):
        return self.__code["realtime"]["best_bid_price"][-1]
    #買進數量
    def get_best_bid_volume(self):
        return self.__code["realtime"]["best_bid_volume"][-1]
    #賣出價格
    def get_best_ask_price(self):
        return self.__code["realtime"]["best_ask_price"][-1]
    #賣出數量
    def get_best_ask_volume(self):
        return self.__code["realtime"]["best_ask_volume"][-1]
    
    #開盤
    def get_open(self):
        return self.__code["realtime"]["open"]
    #高點
    def get_high(self):
        return self.__code["realtime"]["high"]
    #低點
    def get_low(self):
        return self.__code["realtime"]["low"]
    
    #振幅
    def get_amplitude(self):
        return round(((float(get_stock_ing.get_high(self))-float(get_stock_ing.get_low(self)))/float(get_stock_ing.get_low(self)))*100,3)

    def input_data(self):
        # 修改数据
        sheet.range("B2").value = get_stock_ing.get_name(self)
        sheet.range("C2").value = get_stock_ing.get_best_bid_price(self)
        sheet.range("D2").value = get_stock_ing.get_best_ask_price(self)
        sheet.range("E2").value = get_stock_ing.get_latest_trade_price(self)
        sheet.range("F2").value ="-"
        sheet.range("G2").value ="-"
        sheet.range("H2").value =get_stock_ing.get_trade_volume(self)
        sheet.range("I2").value =get_stock_ing.get_best_bid_volume(self)
        sheet.range("J2").value =get_stock_ing.get_best_ask_volume(self)
        sheet.range("K2").value =get_stock_ing.get_accumulate_trade_volume(self)
        sheet.range("L2").value =get_stock_ing.get_high(self)
        sheet.range("M2").value =get_stock_ing.get_low(self)
        sheet.range("N2").value =get_stock_ing.get_open(self)
        #自動調整名稱寬度
        sheet.range("B:B").autofit()

        
        # 保存修改（可选）
        workbook.save()
#--------------------------------------------------------------------        
        

#資料內容 ->文字型態    
def ing(data_list:list):
    stock_data = twstock.realtime.get(data_list)

    for stock_code in stock_data:
        #最後一個會是succcess,所以在最後一個前停下來
        if "success" == stock_code:
            break

        stock=get_stock_ing(stock_data[stock_code])
        stock.input_data()
        print(stock.get_realtime())
    
#資料內容 ->文字型態    
def ed(data_list:list):
    
    for stock_code in data_list:
        #最後一個會是succcess,所以在最後一個前停下來
        if "success" == stock_code:
            break
        stock=get_stock_ed(stock_code)
    
        #列印出資料
        print(stock.price())
        #可15秒可以用
        time.sleep(15)




if __name__ == "__main__":

    # 连接到已打开的Excel应用程序
    app = xw.App(visible=True,add_book=False)

    # 获取活动工作簿
    workbook = app.books.open("data.xlsx")

    # 获取活动工作表
    sheet = workbook.sheets.active




    ing(["0050"])
    #ed(["0050"])
    
