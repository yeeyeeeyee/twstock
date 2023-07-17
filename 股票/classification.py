import xlwings as xw




def classification(data_list:list,read_sheet): 
    # 连接到当前活动的Excel应用程序
    app = xw.apps.active

    # 获取活动工作簿
    wb = app.books.active
    my_list = [s.name for s in wb.sheets]
   
    n=1
    for target in data_list:
        n+=1
        for i ,element in enumerate(my_list):
            if target in element:
                print(f"Element '{element}' contains target '{target}' at position {i}")
                # 将第4行的文字向下移动一格
                #如果read_sheet的range`(a{n})不等於wb.sheets[i]的range`('A5')才執行
                if read_sheet.range(f'A{n}').value != wb.sheets[i].range('A5').value:
                    wb.sheets[i].range("4:4").api.Insert()
                #複製
                read_sheet.range(f'A{n}:o{n}').api.Copy(wb.sheets[i].range('A5').api)
                """ 
                #字中對齊
                wb.sheets[i].range('A5:P5').api.HorizontalAlignment = -4108
                wb.sheets[i].range('A5:P5').api.VerticalAlignment  = -4108

                """
                wb.sheets[i].autofit()
        
    # 保存并关闭工作簿
    wb.save() 


if __name__ == '__main__':
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
    workbook, sheet = main("data.xlsx","a")
    target_elements = ['0050', '0052', '0056', '006208', '00679B', '00687B', '00690', '00692', '00701', '00713', '00728', '00731', '00751B', '00773B', '00850', '00878', '00881', '00888', '1232', '2308', '2317', '2480', '2912', '3711', '8926']
    classification(target_elements,sheet)   
    