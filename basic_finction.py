"""
说明：
基础功能逻辑：读取config.json文件获取常用参数，获取机器人的tenant_access_token，获取用户或机器人所在的群列表chat_id
1、发送文本消息步骤：基础功能，发送文本消息
2、发送图片消息步骤：基础功能，上传图片获取image_key，发送图片
3、发送文件消息步骤：基础功能，上传文件获取file_key，发送文件
"""
import json
import time
from pathlib import Path
import requests
from requests_toolbelt import MultipartEncoder
import os
from openpyxl import load_workbook


# 读取json文件数据，并将其转化为Python字典数据供使用
def read_json_file(json_file_path):
    # 上下文管理器：执行完毕后会自动清理资源，打开后自动关闭
    with open(json_file_path, 'r', encoding='UTF-8-sig') as file:
        # 从文件中读取JSON文件，并将其转化为Python字典
        data = json.load(file)
        return data
    # 文件在这里自动关闭


config_file_path = "config.json"  # 配置文件地址，不特殊说明时，Python程序默认访问的文件路径是相对于当前工作目录的，即程序执行时所处的目录
json_data = read_json_file(config_file_path)  # 读取配置文件
robot_proxies = json_data["configurations"]["robot_proxies"]  # 代理设置
robot_app_id = json_data["configurations"]["App_ID"]  # 飞书开放平台应用的唯一ID标识
robot_app_secret = json_data["configurations"]["App_Secret"]  # 应用的密钥
robot_get_tenant_access_token_url = json_data["configurations"][
    "robot_get_tenant_access_token_url"]  # 获取tenant_access_token的URL
robot_uploadimage_get_image_key_url = json_data["configurations"][
    "robot_uploadimage_get_image_key_url"]  # 上传图片获取image_key的URL
robot_send_info_url = json_data["configurations"]["robot_send_info_url"]  # 发送文本和图片消息的URL
robot_uploadfile_get_file_key_url = json_data["configurations"][
    "robot_uploadfile_get_file_key_url"]  # 上传文件获取file_key的URL


# 自建应用获取 tenant_access_token
def get_tenant_access_token() -> str:
    # 获取tenant_access_token的URL
    url = robot_get_tenant_access_token_url

    # 请求头
    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {
        # app_id和app_secret是自建应用机器人的标识，该机器人是正式版，并已申请获取与上传图片或文件资源的权限
        "app_id": robot_app_id,
        "app_secret": robot_app_secret
    }

    with requests.post(url=url, headers=headers, data=json.dumps(data), proxies=robot_proxies, timeout=10) as response:
        # 发送POST请求，建立连接
        # 由于请求头是"application/json; charset=utf-8"，所以向服务器发送的数据是json格式，那么需要使用json.dumps()进行转换

        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            # 解析json格式的相应内容，响应内容是json格式，需将其转换成Python的字典格式
            token_code = response.json()
            # 返回tenant_access_token值
            return token_code["tenant_access_token"]
        else:
            return "请求失败，状态码：", str(response.status_code)


# 获取用户或机器人所在的群列表，目前是获取小机器人群的chat_id
def get_chat_id(chat_name: str) -> str:
    url = "https://open.feishu.cn/open-apis/im/v1/chats"
    headers = {
        "Authorization": "Bearer " + get_tenant_access_token()
    }

    with requests.get(url=url, headers=headers, proxies=robot_proxies) as response:
        return_data: dict = response.json()

        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            for r in return_data["data"]["items"]:
                if r["name"] == chat_name:
                    chat_id = r["chat_id"]
                    break
            return chat_id
        else:
            return "请求失败，状态码：", str(response.status_code)


# 发送文本消息
def send_text(msg: str, chat_name: str):
    # 查询参数，确定接收者的类型，发送至飞书群
    params = {"receive_id_type": "chat_id"}
    msgcontent = {
        "text": msg,
    }
    # 请求体
    req = {
        "receive_id": get_chat_id(chat_name),  # chat id
        "msg_type": "text",
        "content": json.dumps(msgcontent)
    }
    headers = {
        'Authorization': 'Bearer ' + get_tenant_access_token(),  # your access token
        'Content-Type': 'application/json'
    }

    with requests.post(url=robot_send_info_url, params=params, headers=headers, data=json.dumps(req),
                       proxies=robot_proxies) as response:
        # 异常响应码处理
        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            return "文本发送成功！"
        else:
            return "请求失败，状态码：", response.status_code


# 每隔n秒发送一次文本消息
def timer(n):
    while True:
        send_text("测试数据")
        time.sleep(n)


# 获取图片的image_key
def uploadimage(path, image_name):
    form = {'image_type': 'message',
            'image': (open(path + '/' + image_name, 'rb'))}  # 需要替换具体的path
    # post请求的数据是发送混合数据，使用MultipartEncoder()进行数据传递
    multi_form = MultipartEncoder(form)
    headers = {
        # 获取tenant_access_token, 需要替换为实际的token
        'Authorization': 'Bearer ' + get_tenant_access_token(),
        # 获取post请求数据的content_type，根据数据类型自动改变，无需手动构造
        'Content-Type': multi_form.content_type
    }

    with requests.post(url=robot_uploadimage_get_image_key_url, headers=headers, data=multi_form,
                       proxies=robot_proxies) as response:
        # 异常响应码处理
        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            # 解析json格式的相应内容，响应内容是json格式，需将其转换成Python的字典格式
            token_data = response.json()
            return token_data["data"]["image_key"]
        else:
            return "请求失败，状态码：", response.status_code


# 发送图片
def send_image(image_path: str, image_name: str, chat_name: str):
    # 查询参数，确定接收者的类型，发送至飞书群，需要加在url地址后
    params = {"receive_id_type": "chat_id"}
    # 请求头
    headers = {
        'Authorization': 'Bearer ' + get_tenant_access_token(),  # your access token
        'Content-Type': 'application/json'
    }

    # 图片的image_key的字典类型数据
    image_key_dict = {
        "image_key": uploadimage(image_path, image_name),
    }
    # 请求体
    datas = {
        # 接收者的ID，与接收者类型相同
        "receive_id": get_chat_id(chat_name),   # chat_id
        # 消息类型
        "msg_type": "image",
        # 消息内容：将内容由字典类型转换成json格式（字符串）
        "content": json.dumps(image_key_dict)
    }

    with requests.post(url=robot_send_info_url, params=params, data=json.dumps(datas), proxies=robot_proxies,
                       headers=headers) as response:
        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            return '图片发送成功'
        else:
            return "请求失败，状态码是：" + str(response.status_code)


# 上传文件获取file_key，下面代码中以.xlsx文件为例
def uploadfile(filepath: str, file_name: str):
    form = {'file_type': 'stream',  # 无需修改
            'file_name': file_name,
            'file': (file_name, open(filepath + '/' + file_name, 'rb'),
                     'application/vnd.ms-excel'   # 该参数为根据文件类型进行替换，具体的格式参考  https://www.w3school.com.cn/media/media_mimeref.asp
                     )
            }
    # post请求的数据是发送混合数据，使用MultipartEncoder()进行数据传递
    multi_form = MultipartEncoder(form)
    headers = {
        'Authorization': 'Bearer ' + get_tenant_access_token(),
        # 获取post请求数据的content_type，根据数据类型自动改变，无需手动构造
        'Content-Type': multi_form.content_type
    }

    with requests.post(robot_uploadfile_get_file_key_url, headers=headers, data=multi_form, proxies=robot_proxies) as response:
        # 异常响应码处理
        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            # 解析json格式的相应内容，响应内容是json格式，需将其转换成Python的字典格式
            token_data = response.json()
            return token_data["data"]["file_key"]
        else:
            return "请求失败，状态码：", response.status_code


# 发送文件
def send_file(filepath, file_name):
    # 查询参数，确定接收者的类型，发送至飞书群，需要加在url地址后
    params = {"receive_id_type": "chat_id"}
    # 请求头
    headers = {
        'Authorization': 'Bearer ' + get_tenant_access_token(),  # your access token
        'Content-Type': 'application/json'
    }
    file_key_dict = {
        "file_key": uploadfile(filepath, file_name)
    }
    # 请求体
    data = {
        # 接收者的ID，与接收者类型相同
        "receive_id": get_chat_id("小机器人测试群"),   # chat_id
        # 消息内容：将内容由字典类型转换成json格式（字符串）
        "content": json.dumps(file_key_dict),
        # 消息类型
        "msg_type": "file"
    }
    with requests.post(url=robot_send_info_url, proxies=robot_proxies, headers=headers, data=json.dumps(data),params=params) as response :
        if response.status_code == 200: # 200是状态成功码，请求成功，服务器返回请求的网页
            return '文件发送成功'
        else:
            return "请求失败，状态码是：" + str(response.status_code)


# 读取excel文件，获取文字点评并发送至飞书群
def read_excel_file_get_send_text(excel_file_path, excel_file):

    # excel文件完整路径
    file_path = os.path.join(excel_file_path, excel_file)

    workbook = load_workbook(filename=file_path, data_only=True)
    sheet = workbook.active  # 假设我们读取的是活动工作表

    uppercase_letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    letters = uppercase_letters[0]  # 只取A （取A-D，uppercase_letters[0:4]，不包含终点坐标对应的值）
    num = range(1, 16)   # 不包含终点的数据

    # 使用两层嵌套的列表推导式来生成所有可能的坐标组合，外层循环遍历字母，内层循环遍历数字
    cell_addr = [f"{f}{n}" for f in letters for n in num]
    # 定义空字符串，用来接收单元格对应的值，每接收一个值就存入新的一行
    str_append = ""

    # 遍历所有单元格读取文本，并且返回单元格的值
    for cell in cell_addr:
        cell_value = sheet[cell].value
        str_append += f"{(cell_value or '')}\n"

    # 删除字符串中多余的&
    str_text = str_append.replace("&", "")

    send_text(str_text)




if __name__ == '__main__':
    # 测试获取tenant_access_token
    # tenant_access_token=get_tenant_access_token()
    # print(tenant_access_token)

    # 获取指定群的chat_id
    print(get_chat_id("测试测试"))

    # 发送文字消息，用\n来表示换行
    # send_text("firstline \nsecondline ")

    # 循环发送文字消息
    # timer(5)

    # 发送当前文件夹下的图片，str(Path.cwd())获取当前工作目录
    # send_image(str(Path.cwd()), "图片.png")

    # 发送excel文件
    # send_file(r"D:\MyDocuments\wangjia93\桌面", "2222.xlsx")

    # read_excel_file_get_send_text(r"D:\MyDocuments\wangjia93\桌面\测试文件", "评论.xlsx")