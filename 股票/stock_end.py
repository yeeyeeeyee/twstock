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
            self.get_change(),
            self.get_amplitude()
        ]
        range_address = f"F{row}:G{row}"
        sheet.range(range_address).value = data


#收盤時抓
def update_endofday_data(codes,sheet):
    row=2
    for stock_code in codes:
        if stock_code == "success":
            break
        #如果有錯就下一個
        try:
            stock = StockData(stock_code)
            stock.input_data(sheet,row)
            row += 1
            time.sleep(15)
        except:
            continue