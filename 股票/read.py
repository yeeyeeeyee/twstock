import pandas as pd
import get_stock
from datetime import datetime,time
import time as t
import requests
import stock_end
import json
import xlwings as xw
import classification
import error_case

def open_excel_file(file,sheet_name:str=""):
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

def opening_stock(data_list,sheet):
    # 取得現在的時間
    now = datetime.now().time()
    # 設定下午1點半的時間
    closing_time = time(13, 40)

    error_count = 0  # 错误计数器
    while now < closing_time:
        try:
            get_stock.update_realtime_data(data_list,sheet)
            t.sleep(3)
            now = datetime.now().time()
        except requests.exceptions.ConnectionError as e:
            # 處理連接錯誤
            print("連接錯誤:", str(e))
            t.sleep(5)
        except Exception as e:
            error_count += 1  # 错误计数器加一
            if error_count >= 2:
                error_case.error_case(e)  # 达到错误次数上限，关闭程序
            workbook,sheet = get_stock.main(write_file,write_sheet)

def reload_data():
    print("是否重新載入資料 y/n")
    input_str = input("輸入:")
    if input_str.lower()  == "y":
        get_stock.update_realtime_data(data_list,sheet)
        reload_data()
#-------------------------------------------------
# 讀取 JSON 檔案
with open('股票\\setting.json', 'r',encoding='utf-8') as file:
    config = json.load(file)
#讀取
read_file = config['read_file']
read_sheet = config['read_sheet']

#寫入
write_file=config['write_file']
write_sheet=config['write_sheet']
#存檔
save=config['save']
#等待
check_wait=config['check_wait']
ending_wait=config['ending_wait']


# 關閉檔案
file.close()

#-------------------------------------------------------------------------
# 读取Excel文件
df = pd.read_excel(read_file,read_sheet)

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 1].astype(str).tolist()

# 打印列表
print(data_list)


# 另存为新文件
if save:
    import save_as
    save_as.save_as(read_file)

#開啟檔案
workbook,sheet=open_excel_file(write_file,write_sheet)



#資料
try:
    print("讀取資料")
    stock_end.update_data(data_list,sheet)
    print("即時資料-開始:")
except Exception as e:
    error_case.error_case(e)

#盤中
opening_stock(data_list,sheet)

try:
    #最後補齊即時資料
    get_stock.update_realtime_data(data_list,sheet)

    if check_wait:
        reload_data()

    print("即時資料-結束")

    print("資料分類-開始")
    workbook,sheet=open_excel_file(read_file,read_sheet)
    classification.classification(data_list,sheet)
    print("資料分類-結束")
    if ending_wait:
        input("------請按任意鍵結束-------")
except Exception as e:
    error_case.error_case(e)

#-------------------------------------------------------------------------
 #測試用

#詳細資料
#stock_end.update_data(data_list,sheet)
#即時資料
#get_stock.update_realtime_data(data_list,sheet)
#分類
#workbook,sheet=open_excel_file(read_file,read_sheet)
#classification.classification(data_list,sheet)




