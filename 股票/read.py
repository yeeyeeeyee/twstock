import pandas as pd
import get_stock
from datetime import datetime,time
import time as t
import requests
import stock_end


 
# 读取Excel文件
df = pd.read_excel('88.xlsx')

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 0].astype(str).tolist()

# 打印列表
print(data_list)

# 取得現在的時間
now = datetime.now().time()
# 設定下午1點半的時間
closing_time = time(13, 35)
file="data.xlsx"
sheet_name=""
workbook,sheet = get_stock.main(file,sheet_name)

#get_stock.update_realtime_data(data_list,sheet)
#get_stock.update_endofday_data(data_list,sheet)


error_count = 0  # 错误计数器
#測試5秒會不會被封 答案是不會
while now < closing_time:
    try:
        get_stock.update_realtime_data(data_list,sheet)
        t.sleep(5)
        now = datetime.now().time()
    except requests.exceptions.ConnectionError as e:
        # 處理連接錯誤
        print("連接錯誤:", str(e))
        t.sleep(5)
    except:
        error_count += 1  # 错误计数器加一
        if error_count >= 2:
            exit()  # 达到错误次数上限，关闭程序
        workbook,sheet = get_stock.main(file,sheet_name)

""" 
#1:50分才有漲幅資料
print("暫停中") 
t.sleep(900)           
        
if now > closing_time:
    stock_end.update_endofday_data(data_list,sheet)   
 """



