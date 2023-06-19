import pandas as pd
import get_stock
import time


 
# 读取Excel文件
df = pd.read_excel('88.xlsx')

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 0].astype(str).tolist()

# 打印列表
print(data_list)


#get_stock.ed(data_list)
#get_stock.get_list(data_list)


'''
測試5秒會不會被封 答案是不會
count=0
while 1:
    try:
        count+=1
        get_stock.get_list(data_list)
        print(f'------第{count}次,完后5秒--------')
        time.sleep(5)
    except:
        break
'''



