import requests
from bs4 import BeautifulSoup
import time

def close(soup):
    #獲取成交,開盤  等等資料
    close=soup.find("span",class_="Fw(600) Fz(16px)--mobile Fz(14px) D(f) Ai(c)")
    print(close.text)

   



url = f"https://tw.stock.yahoo.com/quote/0050.TW"
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")
span_elements = soup.find_all("title")
    

close(soup)
    
