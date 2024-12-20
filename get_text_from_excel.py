"""从excel中获取文本"""
import os
import basic_finction
import win32com.client as win32



def initializer_wps_excel():
    """初始化WPS Excel应用程序"""
    try:
        excel = win32.gencache.EnsureDispatch('Ket.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel
    except Exception as e:
        print(f"初始化WPS Excel程序出错，错误为：{e}")
        raise


def open_workbook(excel, excel_path):
    """打开指定路径的工作簿"""
    try:
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))    # 打开指定路径的Excel文件
        return workbook
    except Exception as e:
        print(f"打开工作簿失败: {e}")
        raise


def read_cell_value_to_str(worksheet):
    """依次读取A列单元格内容，拼接成文本"""
    column_text = ""
    row_count = worksheet.UsedRange.Rows.Count    # 获取工作表中包含数据的行数

    try:
        for i in range(1, row_count + 1):     # 从第一行开始遍历，到第row行
            cell_value = worksheet.Cells(i, 1).Value
            # 获取单元格Ai的值，excel.Workbooks.Worksheets.Cells(a,b) 是用来引用 Excel 工作簿中特定工作表的特定单元格的方法。
            # 这里，a 和 b 分别代表行号和列号，它们都是从 1 开始计数的（而不是从 0 开始，这与 Python 的索引不同）。
            if cell_value is not None:
                column_text += str(cell_value) + "\n"  # 将每个单元格的文本拼接，并添加换行符
        return column_text
    except Exception as e:
        print(f"读取单元格值时发生错误:{e}")
        raise


def close_workbook(workbook):
    """关闭工作簿"""
    try:
        workbook.Close(SaveChanges=False)   # 关闭工作簿，不保存更改
    except Exception as e:
        print(f"关闭工作簿失败: {e}")  # 退出WPS Excel应用程序
        raise


def quit_wps_excel(excel):
    """退出WPS Excel应用程序"""
    try:
        excel.Quit()
    except Exception as e:
        print(f"退出WPS Excel失败：{e}")


def capture_excel_cola_data_as_text(excel_path):
    """依次获取A列单元格文本并拼接成文本"""
    excel = initializer_wps_excel()
    try:
        workbook = open_workbook(excel, excel_path)
        worksheet = workbook.Worksheets(1)       # 选择工作簿中的第一个工作表
        column_text = read_cell_value_to_str(worksheet)
        close_workbook(workbook)
        quit_wps_excel(excel)
        return column_text
    except Exception as e:
        print(f"获取A列文本失败：{e}")
        raise


if __name__ == '__main__':
    # excel_file_path = r"D:\MyDocuments\wangjia93\桌面\测试文件"
    excel_file_path = r'C:\Users\Administrator\Desktop\测试文件'
    excel_file_name = "评论.xlsx"
    excel_path = os.path.join(excel_file_path, excel_file_name)
    text = capture_excel_cola_data_as_text(excel_path).replace('&', '')
    basic_finction.send_text(text, "小机器人测试群")









