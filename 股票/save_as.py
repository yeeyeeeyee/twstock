import os
import xlwings as xw
from datetime import datetime
def save_as(file_name):
    # 获取当前时间并格式化为字符串
    current_time = datetime.now().strftime("%Y-%m-%d_%Hh-%Mm-%Ss")
    new_file_name = f"{current_time}.xlsx"

    # 获取当前脚本所在的文件夹路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 构建新文件路径
    new_folder_path = os.path.join(script_dir, 'save_as')
    new_file_path = os.path.join(new_folder_path, new_file_name)
    # 检查新文件路合法性
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)


    # 打开现有工作簿
    wb = xw.Book(file_name)
    # 另存为新文件
    wb.save(new_file_path)
    # 关闭工作簿
    wb.close()

if __name__ == '__main__':
    save_as('99.xlsx')
