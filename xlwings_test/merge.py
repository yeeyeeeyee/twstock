import xlwings as xw
import pandas as pd
#  √  move   ,copy     ,


# 连接到当前活动的Excel应用程序
app = xw.apps.active

# 获取活动工作簿
wb = app.books.active
my_list = [s.name for s in wb.sheets]

# 读取Excel文件
df = pd.read_excel("data.xlsx","a")

# 获取第一列数据并转换为字符串列表
data_list = df.iloc[:, 1].astype(str).tolist()
# 获取A工作表和B工作表
original_sheet = wb.sheets['a']

# 创建字典以加快查找
data_dict = {target: True for target in data_list}
n=1
for target in data_list:
    n+=1
    for i ,element in enumerate(my_list):
        if target in element:
            print(f"Element '{element}' contains target '{target}' at position {i}")
            #print(wb.sheets[i].name)
            wb.sheets[i].range("4:4").api.Insert()
            original_sheet.range(f'B{n}:N{n}').api.Copy(wb.sheets[i].range('B5').api)
# 保存并关闭工作簿
wb.save() 

""" 
# 获取A工作表和B工作表
original_sheet = wb.sheets['a']

# 将第4行的文字向下移动一格
sheet_b.range("4:4").api.Insert()
# 复制B2到N2的范围到B工作表的第4行
sheet_a.range('B2:N2').api.Copy(sheet_b.range('B5').api)



 """
#print(data_list)
#print(my_list)