import twstock
import xlwings as xw
import requests
from bs4 import BeautifulSoup


class RealtimeStockData:
    def __init__(self, code, row):
        self.code = code
        self.row = row
    #       獲取盤中資料
    def get_all_information(self):
        return self.code
    #獲得info裡面個別資料
    def get_info(self):
        return self.code["info"]
    #獲得時間
    #回傳格式  ('2023-06-14', '14:30:00')
    def get_time(self):
        time = self.get_info()["time"].split(" ")
        return time[0]
     #獲得代號
    def get_code(self):
        return self.get_info()["code"]
    #獲得名稱
    def get_name(self):
        return self.get_info()["name"]
    #獲得realtime裡面個別資料
    def get_realtime(self):
        return self.code["realtime"]
    trade_price=""
    #成交價
    def get_latest_trade_price(self):
        if self.get_realtime()["latest_trade_price"] != "-":
            RealtimeStockData.trade_price = self.get_realtime()["latest_trade_price"]
            return self.get_realtime()["latest_trade_price"]
        else:
            return RealtimeStockData.trade_price
 
    #昨收
    #判斷
    def close(self):
        url = f"https://tw.stock.yahoo.com/quote/{self.get_code()}.TW"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        span_elements = soup.find_all("title")
        #如果tw找不到就換TWO
        if span_elements == []:
            url = f"https://tw.stock.yahoo.com/quote/{self.get_code()}.TWO"
            response = requests.get(url)
            soup = BeautifulSoup(response.text, "html.parser")
            span_elements = soup.find_all("title")
        ul=soup.find("ul",class_="D(f) Fld(c) Flw(w) H(192px) Mx(-16px)")
        #變成字典
        dictionary = {key.text: value.text for key, value in ul}
        return dictionary['昨收']
    
    #漲跌
    def get_amplitude(self):
        amplitude=float(self.get_latest_trade_price())-float(RealtimeStockData.close(self))
        return amplitude
    
    # 漲跌%
    def get_amplitude_percent(self):
        amplitude_percent=float(self.get_amplitude())/float(RealtimeStockData.close(self))*100
        amplitude_percent=round(amplitude_percent,2)
        return amplitude_percent
    
    #成交量
    def get_trade_volume(self):
        return self.get_realtime()["trade_volume"]
    #總成交量
    def get_accumulate_trade_volume(self):
        return self.get_realtime()["accumulate_trade_volume"]
    #買進價格
    def get_best_bid_price(self):
        return self.get_realtime()["best_bid_price"][-1]
    #買進數量
    def get_best_bid_volume(self):
        return self.get_realtime()["best_bid_volume"][-1]
    #賣出價格
    def get_best_ask_price(self):
        return self.get_realtime()["best_ask_price"][-1]
    #賣出數量
    def get_best_ask_volume(self):
        return self.get_realtime()["best_ask_volume"][-1]
    #開盤
    def get_open(self):
        return self.get_realtime()["open"]
    #高點
    def get_high(self):
        return self.get_realtime()["high"]
    #低點
    def get_low(self):
        return self.get_realtime()["low"]
    

    #填入資料
    def input_data(self, sheet):
        # 修改数据
        sheet.range(f"A{self.row}").api.NumberFormat = "yyyy/mm/dd" 
        sheet.range(f"A{self.row}").value = self.get_time()
        data = [
            self.get_name(),
            self.get_best_bid_price(),
            self.get_best_ask_price(),
            self.get_latest_trade_price(),
            self.get_amplitude(),
            self.get_amplitude_percent(),
            self.get_trade_volume(),
            self.get_best_bid_volume(),
            self.get_best_ask_volume(),
            self.get_accumulate_trade_volume(),
            self.get_high(),
            self.get_low(),
            self.get_open()
        ]
        #設置c到o
        range_address = f"C{self.row}:O{self.row}"
        #從設置的填入資料
        sheet.range(range_address).value = data
        #自動調整名稱寬度
        sheet.autofit()
        

#已fun的方式來使用realtime.get
def get_stock_data(stock_codes):
    stock_data = twstock.realtime.get(stock_codes)
    return stock_data
#盤中抓即時資料
def update_realtime_data(codes,sheet):
    stock_data = get_stock_data(codes)
    row = 2
    for stock_code, data in stock_data.items():
        if stock_code == "success" :
            break
        #確認列表中有抓不到的資料
        #print(f"stock_code:{stock_code}---data:{data['success']}")

        stock = RealtimeStockData(data, row)
        stock.input_data(sheet)
        row += 1
    # 保存修改
    sheet.book.save()


def main(file,sheet_name:str=""):
    try:
        workbook = xw.Book(file)
    except:
        app = xw.App(visible=True, add_book=False)
        workbook = app.books.open(file)
    if sheet_name == "":
        sheet = workbook.sheets.active
    else:
        sheet = workbook.sheets[sheet_name]
    return workbook, sheet

if __name__ == "__main__":
    workbook, sheet = main("data.xlsx")
    update_realtime_data(["0050","0052"],sheet)
    
