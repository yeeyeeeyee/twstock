import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw



#=============================================================
class end:
    def __init__(self,code):
        self.code=code
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
            self.PE()
            self.PB()
            self.杜邦分析()
            self.NAVPS(soup)
            self.三率()
            self.流速動比率()
            self.負債比()
            self.利息保障倍數()
            self.盈餘再投資比()


            self.HighsAndLows(yahoo)
            self.現金股利發放日_person(soup)
            self.財務報表()
        #ETF
        else:
            self.ManagementFee(elements)
            self.現金股利發放日_ETF(soup)
            self.財務報表()
            self.HighsAndLows(yahoo)

        #---------------------------------------
        
        print(span_elements)
        print(url)
        print("\n")
        

    #漲跌
    def HighsAndLows(self,soup):
        #獲取成交,開盤  等等資料
        ul=soup.find("ul",class_="D(f) Fld(c) Flw(w) H(192px) Mx(-16px)")
        #變成字典
        dictionary = {key.text: value.text for key, value in ul}
        #print(dictionary)
        #獲取字典裡的"昨收"
        print(f"昨收:{dictionary['昨收']}")
        

        
        

    #管理費
    def ManagementFee(self,elements):
        second_element = elements[0]
        desired_elements = second_element.find_all("div",class_="D(f) Ai(fs) H(100%) Fz(16px) Bxz(bb) Bdbw(1px) Bdbs(s) Bdc($bd-primary-divider) Lh(1.5)")
        #所有資料
        pairs = [(element.text, v.text) for element, v in desired_elements]
        #選擇 管理費的值
        print(f"管理費:{pairs[-1][1]}")
        
        

    def 現金股利發放日_ETF(self,soup):
        elements =soup.find_all("div",class_="table-grid Mb(20px) row-fit-half")

        second_element=elements[0]
        desired_elements=second_element.find_all("div",class_="Py(8px) Pstart(12px) Bxz(bb)")
        print(f'現金股利發放日:{desired_elements[-1].text}')
    
        
    #-------------------------------------------------------------------------------------------
    def 現金股利發放日_person(self,soup):
        desired_elements=soup.find_all("div",class_="D(f) Ai(fs) H(100%) Fz(16px) Bxz(bb) Bdbw(1px) Bdbs(s) Bdc($bd-primary-divider) Lh(1.5)")
        dictionary = {key.text: value.text for key, value in desired_elements}
        print(f'現金股利發放日:{dictionary["現金股利發放日"]}')   
        

    #市盈率(PE)
    def PE(self):
        url = f"https://histock.tw/stock/{self.code}/%E6%9C%AC%E7%9B%8A%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        span_elements = soup.find("td", attrs={"style": True})
        print(f"市盈率:{span_elements.text}")
        

    #市淨率
    def PB(self):
        url = f"https://histock.tw/stock/{self.code}/%E8%82%A1%E5%83%B9%E6%B7%A8%E5%80%BC%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        span_elements = soup.find("td", attrs={"style": True})
        print(f"市淨率:{span_elements.text}")
        

    def 財務報表(self):
        url = f"https://histock.tw/stock/{self.code}/%E9%99%A4%E6%AC%8A%E9%99%A4%E6%81%AF"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #除權日
        print(f'除權日:{elements[2].text}')
        #除息日
        print(f'除息日:{elements[1].text}/{elements[3].text}')
        #股票股利
        print(f'股票股利:{elements[5].text}')
        #現金股利
        print(f'現金股利:{elements[6].text}')
        #EPS(盈餘)
        print(f'EPS:{elements[7].text}')
        #現金殖利率(殖利率)
        print(f'現金殖利率:{elements[9].text}')
        

    def 杜邦分析(self):
        url = f"https://histock.tw/stock/{self.code}/%E6%9D%9C%E9%82%A6%E5%88%86%E6%9E%90"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #ROE
        print(f"ROE:{elements[1].text}")
        #稅後淨利率
        print(f"稅後淨利率:{elements[2].text}")
        

    #每股淨值
    def NAVPS(self,soup):
        elements =soup.find("div",class_="table-grid Mb(20px) row-fit-half", attrs={"style": True})
        div=elements.find_all("div",class_="D(f) Ai(fs) H(100%) Fz(16px) Bxz(bb) Bdbw(1px) Bdbs(s) Bdc($bd-primary-divider) Lh(1.5)")

        dictionary = {key.text: value.text for key, value in div}
        print(f'每股淨值:{dictionary["每股淨值"]}')
        

    def 三率(self):
        url = f"https://histock.tw/stock/{self.code}/%E5%88%A9%E6%BD%A4%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #毛利率
        print(f"毛利率:{elements[1].text}")
        #營益率
        print(f"營益率:{elements[2].text}")
        #淨利率
        print(f"淨利率:{elements[4].text}")
        

    def 流速動比率(self):
        url = f"https://histock.tw/stock/{self.code}/%E6%B5%81%E9%80%9F%E5%8B%95%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #流動比
        print(f"流動比:{elements[1].text}")
        #速動比
        print(f"速動比:{elements[2].text}")
        

    def 負債比(self):
        url = f"https://histock.tw/stock/{self.code}/%E8%B2%A0%E5%82%B5%E4%BD%94%E8%B3%87%E7%94%A2%E6%AF%94"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #負債比
        print(f"負債比:{elements[1].text}")
        

    def 利息保障倍數(self):
        url = f"https://histock.tw/stock/{self.code}/%E5%88%A9%E6%81%AF%E4%BF%9D%E9%9A%9C%E5%80%8D%E6%95%B8"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #利息保障倍數
        print(f"利息保障倍數:{elements[1].text}")
        

    def 營運週轉天數(self):
        url = f"https://histock.tw/stock/{self.code}/%E7%87%9F%E9%81%8B%E9%80%B1%E8%BD%89%E5%A4%A9%E6%95%B8"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #應收帳款收現天數
        print(f"應收帳款收現天數:{elements[1].text}")
        #存貨週轉天數
        print(f"存貨週轉天數:{elements[2].text}")
        

    def 盈餘再投資比(self):
        url = f"https://histock.tw/stock/{self.code}/%E7%9B%88%E9%A4%98%E5%86%8D%E6%8A%95%E8%B3%87%E6%AF%94%E7%8E%87"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        elements = soup.find_all("td")
        #盈餘再投資比
        print(f"盈餘再投資比:{elements[1].text}")



#盤中抓即時資料
def update_data(codes:list,sheet):
    stock_data = codes
    row = 2
    for  data in stock_data:
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

if __name__ == '__main__':
  """  
    # 读取Excel文件
    df = pd.read_excel('88.xlsx')

    # 获取第一列数据并转换为字符串列表
    data_list = df.iloc[:, 1].astype(str).tolist()

    #測試用資料
    #data_list = ["0050","2912","1232","0056"]
    # 打印列表
    print(data_list)
    #判斷列表股票代碼資料 
    for i in data_list:
        code=end(i)
        code.judge() """
    
    

   



    
    

   
