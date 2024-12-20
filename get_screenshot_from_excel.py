"""
自动化地从WPS Excel中捕获包含数据的单元格区域，并将其保存为图像文件。
"""
import os
from PIL import ImageGrab
import win32com.client as win32



def initialize_wps_excel():
    """初始化WPS Excel应用程序，打开WPS Excel应用程序"""
    try:
        excel = win32.gencache.EnsureDispatch('Ket.Application')    # 创建一个WPS Excel应用程序的实例。Ket.Application是WPS Excel的ProgID，用于唯一标识WPS Excel应用程序。这个实例代表了正在运行的 WPS Excel 应用程序
        excel.Visible = False    # 设置WPS Excel在后台运行，不显示用户界面
        excel.DisplayAlerts = False # 禁止显示警告信
        return excel
    except Exception as e:
        print(f"初始化WPS Excel失败: {e}")
        raise

def open_workbook(excel, excel_path):
    """打开指定路径的工作簿"""
    try:
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))    # 打开指定路径的Excel文件
        return workbook
    except Exception as e:
        print(f"打开工作簿失败: {e}")
        raise       # raise 关键字确保了即使在异常被捕获并打印错误消息后，异常仍然可以被传播到上层调用者，使得异常可以被进一步处理或记录

def copy_used_range_to_clipboard(worksheet):
    """复制包含数据的单元格区域到剪贴板"""
    try:
        used_range = worksheet.UsedRange  # 获取工作表中包含数据的单元格区域
        used_range.CopyPicture()  # 将包含数据的单元格区域复制为图片
        worksheet.Paste()  # 将复制的图片粘贴到工作表中
        shape = worksheet.Shapes(worksheet.Shapes.Count)  # 获取粘贴后创建的形状（图片）
        shape.Copy()  # 将形状（图片）复制到剪贴板
    except Exception as e:
        print(f"复制单元格区域到剪贴板失败: {e}")
        raise

def save_image_from_clipboard(output_path):
    """从剪贴板中获取图像并保存"""
    try:
        img = ImageGrab.grabclipboard()     # 从剪贴板获取图像
        img.save(os.path.abspath(output_path))  # 将图像保存为文件
    except Exception as e:
        print(f"保存图像失败: {e}")
        raise

def close_workbook(workbook):
    """关闭工作簿"""
    try:
        workbook.Close(SaveChanges=False)   # 关闭工作簿，不保存更改
    except Exception as e:
        print(f"关闭工作簿失败: {e}")  # 退出WPS Excel应用程序
        raise

def quit_excel(excel):
    """退出WPS Excel应用程序"""
    try:
        excel.Quit()
    except Exception as e:
        print(f"退出WPS Excel失败: {e}")
        raise


def capture_excel_data_as_image(excel_path, output_path):
    """捕获Excel数据作为图像"""
    excel = initialize_wps_excel()  # 创建Excel实例
    try:
        workbook = open_workbook(excel, excel_path)     # 打开工作簿
        worksheet = workbook.Worksheets(1)      # 选择工作簿中的第一个工作表
        copy_used_range_to_clipboard(worksheet)
        save_image_from_clipboard(output_path)
    finally:
        close_workbook(workbook)
        quit_excel(excel)


if __name__ == '__main__':
    excel_file_path = r"D:\MyDocuments\wangjia93\桌面\测试文件"
    excel_file_name = "config.xlsx"
    excel_path = os.path.join(excel_file_path, excel_file_name)
    # 截图文件存放路径
    output_image_path = os.path.join(excel_file_path, '截图.png')
    # 捕获Excel数据作为图像
    capture_excel_data_as_image(excel_path, output_image_path)
