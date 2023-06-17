import xlwings as xw

# 检查 Excel 文件是否已经打开
def is_excel_open(filename):
    for app in xw.apps:
        for book in app.books:
            if book.name == filename:
                workbook = app.books[filename]
                #app.quit()
                return True
    return False

# 打开 Excel 文件
def open_excel(filename):
    app = xw.App(visible=True,add_book=False) 
    workbook = app.books.open(filename)  # 打开指定文件
    return app, workbook

# 主函数
def main():
    filename = 'data.xlsx'  # Excel 文件名

    if is_excel_open(filename):
        print(f"{filename} 文件已经是打开的。")
        
    else:
        app, book = open_excel(filename)
        print(f"{filename} 文件打开。")

if __name__ == '__main__':
    main()
    
