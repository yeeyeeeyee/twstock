import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw



#=============================================================
class end:
    def __init__(self,code,row):
        self.code=code
        self.row=row
        self.昨收 = "-"
        self.市盈率 = "-"
        self.市淨率 = "-"
        self.ROE = "-"
        self.資產報酬率 = "-"
        self.毛利率 = "-"
        self.營益率 = "-"
        self.稅後淨利率 = "-"
        self.每股淨值 = "-"
        self.盈餘 = "-"
        self.流動比率 = "-"
        self.速動比率 = "-"
        self.負債比率 = "-"
        self.利息保障倍數 = "-"
        self.應收帳款收現天數 = "-"
        self.存貨週轉天數 = "-"
        self.現金股利 = "-"
        self.股票股利 = "-"
        self.殖利率 = "-"
        self.除息日 = "-"
        self.股息發放日 = "-"
        self.除權日 = "-"
        self.盈餘再投資比 = "-"
        self.管理費 = "-"
            
    
        

    #資料
    def yesterday_close(self,soup):
        close=soup.find("span",class_="Fw(600) Fz(16px)--mobile Fz(14px) D(f) Ai(c)")
        self.昨收=close.text
                
    #管理費
    def ManagementFee(self,soup):
        elements =soup.find("div",class_="Py(8px) Pstart(12px) Bxz(bb) etf-management-fee")
        self.管理費=elements.text
        

    def 股息發放日_ETF(self,soup):
        elements =soup.find_all("div",class_="table-grid Mb(20px) row-fit-half")

        second_element=elements[0]
        desired_elements=second_element.find_all("div",class_="Py(8px) Pstart(12px) Bxz(bb)")
        self.股息發放日=desired_elements[-1].text
        #print(f'股息發放日:{desired_elements[-1].text}')
    
        
    #-------------------------------------------------------------------------------------------
    def 股息發放日_person(self,soup):
        elements =soup.find_all("div",class_="table-grid Mb(20px) row-fit-half", attrs={"style": True})
        second_element=elements[1]
        find= second_element.find_all("div",class_="Py(8px) Pstart(12px) Bxz(bb)")
        self.股息發放日=find[-1].text
        
        

    #市盈率(PE)
    def get_PE(self):
        url = f"https://histock.tw/stock/{self.code}/%E6%9C%AC%E7%9B%8A%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        span_elements = soup.find("td", attrs={"style": True})
        self.市盈率=span_elements.text
        #print(f"市盈率:{span_elements.text}")
        

    #市淨率
    def get_PB(self):
        url = f"https://histock.tw/stock/{self.code}/%E8%82%A1%E5%83%B9%E6%B7%A8%E5%80%BC%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        span_elements = soup.find("td", attrs={"style": True})
        self.市淨率=span_elements.text
        #print(f"市淨率:{span_elements.text}")
        

    def 財務報表(self):
        url = f"https://histock.tw/stock/{self.code}/%E9%99%A4%E6%AC%8A%E9%99%A4%E6%81%AF"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #除權日
        self.除權日=elements[2].text
        #print(f'除權日:{elements[2].text}')
        #除息日
        self.除息日=f'{elements[1].text}/{elements[3].text}'
        #print(f'除息日:{elements[1].text}/{elements[3].text}')
        #股票股利
        self.股票股利=elements[5].text
        #print(f'股票股利:{elements[5].text}')
        #現金股利
        self.現金股利=elements[6].text
        #print(f'現金股利:{elements[6].text}')
        #EPS(盈餘)
        self.盈餘=elements[7].text
        #print(f'EPS:{elements[7].text}')
        #現金殖利率(殖利率)
        self.殖利率=elements[9].text
        #print(f'現金殖利率:{elements[9].text}')
        

    def 杜邦分析(self):
        url = f"https://histock.tw/stock/{self.code}/%E5%A0%B1%E9%85%AC%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #ROE
        self.ROE=elements[1].text
        #print(f"ROE:{elements[1].text}")
        #ROA
        self.資產報酬率=elements[2].text
        
        

    #每股淨值
    def NAVPS(self,soup):
        elements =soup.find("div",class_="table-grid Mb(20px) row-fit-half", attrs={"style": True})
        second_element=elements.find_all("div",class_="Py(8px) Pstart(12px) Bxz(bb)")
        self.每股淨值=second_element[-1].text
        
        

    def 三率(self):
        url = f"https://histock.tw/stock/{self.code}/%E5%88%A9%E6%BD%A4%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #毛利率
        self.毛利率=elements[1].text
        #print(f"毛利率:{elements[1].text}")
        #營益率
        self.營益率=elements[2].text
        #print(f"營益率:{elements[2].text}")
        #稅後淨利率
        self.稅後淨利率=elements[4].text
        ##print(f"淨利率:{elements[4].text}")
        

    def 流速動比率(self):
        url = f"https://histock.tw/stock/{self.code}/%E6%B5%81%E9%80%9F%E5%8B%95%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #流動比
        self.流動比率=elements[1].text
        #print(f"流動比:{elements[1].text}")
        #速動比
        self.速動比率=elements[2].text
        #print(f"速動比:{elements[2].text}")
        

    def 負債比(self):
        url = f"https://histock.tw/stock/{self.code}/%E8%B2%A0%E5%82%B5%E4%BD%94%E8%B3%87%E7%94%A2%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #負債比
        self.負債比率=elements[1].text
        #print(f"負債比:{elements[1].text}")
        

    def get_利息保障倍數(self):
        url = f"https://histock.tw/stock/{self.code}/%E5%88%A9%E6%81%AF%E4%BF%9D%E9%9A%9C%E5%80%8D%E6%95%B8"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #利息保障倍數
        self.利息保障倍數=elements[1].text
        #print(f"利息保障倍數:{elements[1].text}")
        

    def 營運週轉天數(self):
        url = f"https://histock.tw/stock/{self.code}/%E7%87%9F%E9%81%8B%E9%80%B1%E8%BD%89%E5%A4%A9%E6%95%B8"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #應收帳款收現天數
        self.應收帳款收現天數=elements[1].text
        #print(f"應收帳款收現天數:{elements[1].text}")
        #存貨週轉天數
        self.存貨週轉天數=elements[2].text
        #print(f"存貨週轉天數:{elements[2].text}")
        

    def get_盈餘再投資比(self):
        url = f"https://histock.tw/stock/{self.code}/%E7%9B%88%E9%A4%98%E5%86%8D%E6%8A%95%E8%B3%87%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #盈餘再投資比
        self.盈餘再投資比=elements[1].text
        #print(f"盈餘再投資比:{elements[1].text}")

    #判斷
    def judge(self):
        url = f"https://tw.stock.yahoo.com/quote/{self.code}.TW"
        response = requests.get(url)
        yahoo = BeautifulSoup(response.text, "html.parser")
        span_elements = yahoo.find_all("title")

        
        #如果tw找不到就換TWO
        if span_elements == []:
            url = f"https://tw.stock.yahoo.com/quote/{self.code}.TWO"
            response = requests.get(url)
            yahoo = BeautifulSoup(response.text, "html.parser")
            span_elements = yahoo.find_all("title")
        

        #判斷是否為個股
        url += "/profile"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        #yahoo重要行事曆
        elements =soup.find_all("div",class_="table-grid row-fit-half", attrs={"style": True})
        #個股
        if len(elements)==4:
            self.get_PE()
            self.get_PB()
            self.杜邦分析()
            self.NAVPS(soup)
            self.三率()
            self.流速動比率()
            self.負債比()
            self.營運週轉天數()
            self.get_利息保障倍數()
            self.get_盈餘再投資比()


            self.yesterday_close(yahoo)
            self.股息發放日_person(soup)
            self.財務報表()
        #ETF
        else:
            self.ManagementFee(soup)
            self.股息發放日_ETF(soup)
            self.財務報表()
            self.yesterday_close(yahoo)

        #---------------------------------------
        
        #print(span_elements)
        #print(url)
        #print("\n")

    def input_data(self,sheet):
        url = f"https://tw.stock.yahoo.com/quote/{self.code}.TW"
        response = requests.get(url)
        yahoo = BeautifulSoup(response.text, "html.parser")
        span_elements = yahoo.find_all("title")

        
        #如果tw找不到就換TWO
        if span_elements == []:
            url = f"https://tw.stock.yahoo.com/quote/{self.code}.TWO"
            response = requests.get(url)
            yahoo = BeautifulSoup(response.text, "html.parser")
            span_elements = yahoo.find_all("title")
        
        print(span_elements)
        #判斷是否為個股
        url += "/profile"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        #yahoo重要行事曆
        elements =soup.find_all("div",class_="table-grid row-fit-half", attrs={"style": True})
        #個股
        if len(elements)==4:
            self.get_PE()
            self.get_PB()
            self.杜邦分析()
            self.NAVPS(soup)
            self.三率()
            self.流速動比率()
            self.負債比()
            self.營運週轉天數()
            self.get_利息保障倍數()
            self.get_盈餘再投資比()


            self.yesterday_close(yahoo)
            self.股息發放日_person(soup)
            self.財務報表()
        #ETF
        else:
            self.ManagementFee(soup)
            self.股息發放日_ETF(soup)
            self.財務報表()
            self.yesterday_close(yahoo)


        data=[
            self.昨收 ,
            self.市盈率 ,
            self.市淨率,
            self.ROE ,
            self.資產報酬率 ,
            self.毛利率 ,
            self.營益率 ,
            self.稅後淨利率 ,
            self.每股淨值 ,
            self.盈餘 ,
            self.流動比率 ,
            self.速動比率 ,
            self.負債比率 ,
            self.利息保障倍數 ,
            self.應收帳款收現天數 ,
            self.存貨週轉天數 ,
            self.現金股利 ,
            self.股票股利 ,
            self.殖利率 ,
            self.除息日 ,
            self.股息發放日 ,
            self.除權日 ,
            self.盈餘再投資比 ,
            self.管理費 ,
        ]
         #設置P到AM
        range_address = f"P{self.row}:AM{self.row}"
        #從設置的填入資料
        sheet.range(range_address).value = data
        #自動調整名稱寬度
        sheet.autofit()



#盤中抓即時資料
def update_data(codes:list,sheet):
    stock_data = codes
    row = 2
    for  data in stock_data:
        stock = end(data, row)
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

if __name__ == '__main__':
  workbook, sheet = main("data.xlsx")
  update_data(["1232","2105","2308"],sheet)

    
    

   



    
    

   
