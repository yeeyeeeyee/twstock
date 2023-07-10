import pandas as pd
import get_stock
from datetime import datetime,time
import time as t
import requests
import stock_end
import json

# 讀取 JSON 檔案
with open('股票\\seting.json', 'r',encoding='utf-8') as file:
    config = json.load(file)
#讀取
read_file = config['read_file']
read_sheet = config['read_sheet']

#寫入
write_file=config['write_file']
write_sheet=config['write_sheet']

# 關閉檔案
file.close()
# 读取Excel文件
df = pd.read_excel(read_file,read_sheet)

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 1].astype(str).tolist()

# 打印列表
print(data_list)

# 取得現在的時間
now = datetime.now().time()
# 設定下午1點半的時間
closing_time = time(13, 40)

workbook,sheet = get_stock.main(write_file,write_sheet)

#測試用

#詳細資料
#stock_end.update_data(data_list,sheet)
#即時資料
#get_stock.update_realtime_data(data_list,sheet)


error_count = 0  # 错误计数器

#資料
print("讀取資料")
stock_end.update_data(data_list,sheet)
print("即時資料-開始:")

while now < closing_time:
#while 1:
    try:
        get_stock.update_realtime_data(data_list,sheet)
        t.sleep(3)
        now = datetime.now().time()
    except requests.exceptions.ConnectionError as e:
        # 處理連接錯誤
        print("連接錯誤:", str(e))
        t.sleep(5)
    except:
        error_count += 1  # 错误计数器加一
        if error_count >= 2:
            exit()  # 达到错误次数上限，关闭程序
        workbook,sheet = get_stock.main(write_file,write_sheet)

#最後補齊即時資料
get_stock.update_realtime_data(data_list,sheet)
print("結束")




