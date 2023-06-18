import twstock
import time
import xlwings as xw

class StockData:
    def __init__(self, code):
        self.code = code
        self.stock = twstock.Stock(code)
        
    def get_info(self):
        return twstock.codes[self.code]

    def get_id(self):
        return self.stock.sid

    def get_date(self):
        return self.stock.date[-1]

    def get_change(self):
        return self.stock.change[-1]

    def get_high(self):
        return self.stock.high[-1]

    def get_low(self):
        return self.stock.low[-1]

    def get_price(self):
        return self.stock.price[-1]

    def get_open(self):
        return self.stock.open[-1]

    def get_close(self):
        return self.stock.close[-1]

    def get_amplitude(self):
        high = self.get_high()
        low = self.get_low()
        return round(((float(high) - float(low)) / float(low)) * 100, 2)

    def input_data(self, sheet, row):
        data = [
            self.code,
            self.get_change(),
            self.get_amplitude()
        ]
        range_address = f"F{row}:G{row}"
        sheet.range(range_address).value = data

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
        return time[0], time[1]
     #獲得代號
    def get_code(self):
        return self.get_info()["code"]
    #獲得名稱
    def get_name(self):
        return self.get_info()["name"]
    #獲得realtime裡面個別資料
    def get_realtime(self):
        return self.code["realtime"]
    #成交價
    def get_latest_trade_price(self):
        return self.get_realtime()["latest_trade_price"]
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
    #振幅
    def get_amplitude(self):
        high = self.get_high()
        low = self.get_low()
        return round(((float(high) - float(low)) / float(low)) * 100, 2)

    #填入資料
    def input_data(self, sheet):
        # 修改数据
        data = [
            self.get_name(),
            self.get_best_bid_price(),
            self.get_best_ask_price(),
            self.get_latest_trade_price(),
            "-",
            "-",
            self.get_trade_volume(),
            self.get_best_bid_volume(),
            self.get_best_ask_volume(),
            self.get_accumulate_trade_volume(),
            self.get_high(),
            self.get_low(),
            self.get_open()
        ]
        #設置b到n
        range_address = f"B{self.row}:N{self.row}"
        #從設置的填入資料
        sheet.range(range_address).value = data
        #自動調整名稱寬度
        sheet.range("B:B").autofit()
        # 保存修改
        sheet.book.save()

#已fun的方式來使用realtime.get
def get_stock_data(stock_codes):
    stock_data = twstock.realtime.get(stock_codes)
    return stock_data
#盤中抓即時資料
def update_realtime_data(codes,sheet):
    stock_data = get_stock_data(codes)
    row = 2
    for stock_code, data in stock_data.items():
        if stock_code == "success":
            break

        stock = RealtimeStockData(data, row)
        stock.input_data(sheet)
        row += 1
#收盤時抓
def update_endofday_data(codes,sheet):
    row=2
    for stock_code in codes:
        if stock_code == "success":
            break
        stock = StockData(stock_code)
        stock.input_data(sheet,row)
        row += 1
        time.sleep(15)

def main():
    try:
        workbook = xw.Book("data.xlsx")
    except:
        app = xw.App(visible=True, add_book=False)
        workbook = app.books.open("data.xlsx")
    
    sheet = workbook.sheets.active
    return workbook, sheet

if __name__ == "__main__":
    workbook, sheet = main()
    update_realtime_data(["0050","0052"],sheet)
    update_endofday_data(["0050","0052"],sheet)
