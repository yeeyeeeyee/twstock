import pandas as pd
import get_stock




# 读取Excel文件
df = pd.read_excel('88.xlsx')

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 0].astype(str).tolist()

# 打印列表
print(data_list)



#get_stock.get_list(data_list)


